"""Construcción de gráficos interactivos y animados de Curva S con Apache ECharts."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path
import pandas as pd


def _load_echarts_js() -> str:
    """Carga el script local de echarts si existe, o enlaza al CDN como fallback."""
    local_path = Path(__file__).resolve().parent / "assets" / "echarts.min.js"
    if local_path.is_file() and local_path.stat().st_size > 100_000:
        try:
            return f"<script>{local_path.read_text(encoding='utf-8')}</script>"
        except Exception:
            pass
    return '<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>'


def build_curva_s_echarts_html(
    df_prog: pd.DataFrame,
    df_real_pyrevit: pd.DataFrame | None = None,
    frequency: str = "W",
    titulo: str = "Curva S · Programado vs. Real",
    cutoff_date: date | None = None,
) -> str:
    """Genera el HTML autónomo con Apache ECharts para la Curva S animada."""
    from logic.charts import _alinear_real_con_semanas, _current_week_label

    df_p = df_prog.copy() if df_prog is not None else pd.DataFrame()

    # Eje X y datos programados
    if frequency == "W":
        x_labels = (
            df_p["semana_label"].astype(str).tolist()
            if "semana_label" in df_p.columns
            else [f"S{i+1:02d}" for i in range(len(df_p))]
        )
        y_prog = [
            round(float(v), 2) if pd.notna(v) else 0.0
            for v in df_p.get("acumulado", [])
        ]
    else:
        fechas = pd.to_datetime(df_p.get("fecha", []), errors="coerce")
        x_labels = [
            d.strftime("%d-%m-%Y") if pd.notna(d) else ""
            for d in fechas
        ]
        y_prog = [
            round(float(v), 2) if pd.notna(v) else 0.0
            for v in df_p.get("acumulado", [])
        ]

    # Datos reales (alineados al mismo eje X)
    has_real = False
    y_real: list[float | None] = []

    if df_real_pyrevit is not None and not df_real_pyrevit.empty and "pct_ejecutado_acum" in df_real_pyrevit.columns:
        if frequency == "W":
            x_real_aligned, y_real_aligned = _alinear_real_con_semanas(df_p, df_real_pyrevit)
            real_map = dict(zip(x_real_aligned, y_real_aligned))
            if real_map:
                has_real = True
                # Rellenar con null los períodos futuros que aún no tienen ejecución
                for lbl in x_labels:
                    val = real_map.get(lbl)
                    y_real.append(round(float(val), 2) if val is not None else None)
        else:
            # Curva diaria
            df_r = df_real_pyrevit.copy()
            df_r["fecha_dt"] = pd.to_datetime(df_r["fecha"], errors="coerce", dayfirst=True)
            df_r = df_r.dropna(subset=["fecha_dt"]).sort_values("fecha_dt")
            real_date_map = {
                row["fecha_dt"].strftime("%d-%m-%Y"): round(float(row["pct_ejecutado_acum"]), 2)
                for _, row in df_r.iterrows()
            }
            if real_date_map:
                has_real = True
                last_real_val = None
                for lbl in x_labels:
                    if lbl in real_date_map:
                        last_real_val = real_date_map[lbl]
                        y_real.append(last_real_val)
                    else:
                        y_real.append(None)

    # Identificar marca vertical de Hoy / Corte
    mark_label = None
    if cutoff_date is not None:
        if frequency == "W":
            mark_label = _current_week_label(df_p, cutoff_date)
        else:
            cutoff_str = cutoff_date.strftime("%d-%m-%Y")
            if cutoff_str in x_labels:
                mark_label = cutoff_str

    echarts_script = _load_echarts_js()

    chart_config = {
        "x_labels": x_labels,
        "y_prog": y_prog,
        "y_real": y_real if has_real else [],
        "has_real": has_real,
        "title": titulo,
        "mark_label": mark_label,
    }

    config_json = json.dumps(chart_config)

    html_code = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    {echarts_script}
    <style>
        html, body {{
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
            background-color: transparent;
            overflow: hidden;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        #chart-container {{
            width: 100%;
            height: 480px;
        }}
    </style>
</head>
<body>
    <div id="chart-container"></div>
    <script>
        (function() {{
            const data = {config_json};
            const dom = document.getElementById('chart-container');
            if (!dom) return;

            const chart = echarts.init(dom, null, {{ renderer: 'canvas' }});

            const seriesList = [
                {{
                    name: 'Programado % (Excel)',
                    type: 'line',
                    smooth: 0.35,
                    showSymbol: false,
                    symbol: 'circle',
                    symbolSize: 6,
                    lineStyle: {{ width: 3.5, color: '#0077B6' }},
                    itemStyle: {{ color: '#0077B6' }},
                    areaStyle: {{
                        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                            {{ offset: 0, color: 'rgba(0, 119, 182, 0.32)' }},
                            {{ offset: 1, color: 'rgba(0, 119, 182, 0.02)' }}
                        ])
                    }},
                    data: data.y_prog
                }}
            ];

            const legendData = ['Programado % (Excel)'];

            if (data.has_real && data.y_real.length > 0) {{
                legendData.push('Real % Ejecutado (BIM)');
                seriesList.push({{
                    name: 'Real % Ejecutado (BIM)',
                    type: 'line',
                    smooth: 0.35,
                    showSymbol: false,
                    symbol: 'rect',
                    symbolSize: 7,
                    lineStyle: {{ width: 3, color: '#00A0E3', type: 'dashed' }},
                    itemStyle: {{ color: '#00A0E3' }},
                    areaStyle: {{
                        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                            {{ offset: 0, color: 'rgba(0, 160, 227, 0.28)' }},
                            {{ offset: 1, color: 'rgba(0, 160, 227, 0.01)' }}
                        ])
                    }},
                    data: data.y_real
                }});
            }}

            // Línea indicadora de fecha de corte
            if (data.mark_label) {{
                seriesList[0].markLine = {{
                    symbol: ['none', 'none'],
                    lineStyle: {{ color: '#EF4444', width: 2, type: 'dashed' }},
                    label: {{
                        show: true,
                        position: 'insideEndTop',
                        formatter: 'Hoy: ' + data.mark_label,
                        color: '#EF4444',
                        fontSize: 11,
                        fontWeight: 'bold',
                        backgroundColor: 'rgba(15, 23, 42, 0.8)',
                        padding: [3, 6],
                        borderRadius: 3
                    }},
                    data: [{{ xAxis: data.mark_label }}]
                }};
            }}

            const option = {{
                backgroundColor: 'transparent',
                animation: true,
                animationDuration: 2200,
                animationEasing: 'cubicInOut',
                title: {{
                    text: data.title,
                    textStyle: {{ color: '#E2E8F0', fontSize: 15, fontWeight: 'bold' }},
                    left: 4,
                    top: 2
                }},
                tooltip: {{
                    trigger: 'axis',
                    backgroundColor: 'rgba(15, 23, 42, 0.94)',
                    borderColor: '#334155',
                    borderWidth: 1,
                    textStyle: {{ color: '#F8FAFC', fontSize: 12 }},
                    axisPointer: {{
                        type: 'cross',
                        lineStyle: {{ color: '#94A3B8', type: 'dashed' }},
                        crossStyle: {{ color: '#94A3B8' }}
                    }},
                    formatter: function(params) {{
                        if (!params || !params.length) return '';
                        let header = '<div style="font-weight:700;margin-bottom:6px;border-bottom:1px solid #334155;padding-bottom:3px;color:#94A3B8;">' + params[0].axisValue + '</div>';
                        let progVal = null;
                        let realVal = null;
                        let rows = '';

                        params.forEach(item => {{
                            if (item.value === null || item.value === undefined) return;
                            const isProg = item.seriesName.includes('Programado');
                            const colorDot = '<span style="display:inline-block;margin-right:7px;border-radius:50%;width:9px;height:9px;background-color:' + item.color + '"></span>';
                            const valStr = Number(item.value).toFixed(1) + '%';

                            rows += '<div style="display:flex;justify-content:space-between;gap:20px;margin:3px 0;">' +
                                   '<span>' + colorDot + item.seriesName + '</span>' +
                                   '<b style="color:' + item.color + ';">' + valStr + '</b></div>';

                            if (isProg) progVal = Number(item.value);
                            else realVal = Number(item.value);
                        }});

                        if (progVal !== null && realVal !== null) {{
                            const diff = realVal - progVal;
                            const isAhead = diff >= 0;
                            const diffColor = isAhead ? '#10B981' : '#EF4444';
                            const diffIcon = isAhead ? '▲ Adelanto' : '▼ Atraso';
                            rows += '<div style="margin-top:6px;padding-top:5px;border-top:1px dashed #334155;color:' + diffColor + ';font-weight:700;display:flex;justify-content:space-between;">' +
                                   '<span>' + diffIcon + ':</span><span>' + (isAhead ? '+' : '') + diff.toFixed(1) + '%</span></div>';
                        }}
                        return header + rows;
                    }}
                }},
                legend: {{
                    data: legendData,
                    top: 6,
                    right: 8,
                    textStyle: {{ color: '#94A3B8', fontSize: 11 }},
                    itemGap: 14
                }},
                grid: {{
                    left: '1%',
                    right: '3%',
                    bottom: '12%',
                    top: '16%',
                    containLabel: true
                }},
                xAxis: {{
                    type: 'category',
                    data: data.x_labels,
                    boundaryGap: false,
                    axisLine: {{ lineStyle: {{ color: '#334155' }} }},
                    axisLabel: {{ color: '#94A3B8', fontSize: 10, rotate: 30 }},
                    splitLine: {{ show: true, lineStyle: {{ color: 'rgba(51, 65, 85, 0.4)', type: 'dashed' }} }}
                }},
                yAxis: {{
                    type: 'value',
                    min: 0,
                    max: 110,
                    axisLabel: {{
                        color: '#94A3B8',
                        fontSize: 11,
                        formatter: '{{value}}%'
                    }},
                    splitLine: {{ lineStyle: {{ color: 'rgba(51, 65, 85, 0.45)', type: 'dashed' }} }}
                }},
                dataZoom: [
                    {{
                        type: 'inside',
                        start: 0,
                        end: 100
                    }},
                    {{
                        type: 'slider',
                        start: 0,
                        end: 100,
                        height: 18,
                        bottom: 0,
                        borderColor: '#1E293B',
                        backgroundColor: 'rgba(15, 23, 42, 0.6)',
                        fillerColor: 'rgba(0, 119, 182, 0.25)',
                        handleStyle: {{ color: '#0077B6', borderColor: '#38BDF8' }},
                        textStyle: {{ color: '#64748B', fontSize: 10 }}
                    }}
                ],
                series: seriesList
            }};

            chart.setOption(option);
            window.addEventListener('resize', function() {{
                chart.resize();
            }});
        }})();
    </script>
</body>
</html>"""
    return html_code
