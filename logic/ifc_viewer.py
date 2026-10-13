"""Carga de propiedades y geometría de modelos IFC para el dashboard."""

from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path

import ifcopenshell
import ifcopenshell.geom
import pandas as pd
from ifcopenshell.util.element import get_psets


def _normalize_name(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(character for character in text if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]", "", text.casefold())


PROPERTY_ALIASES = {
    "weight": {"weigth", "weight"},
    "assembly_code": {"codigodemontaje", "assemblycode", "assembly_code"},
    "assembly_description": {
        "descripciondemontaje", "assemblydescription"
    },
    "ei_zona": {"eizona"},
    "ei_subcontratista": {"eisubcontratista"},
    "ei_semana": {"e semana", "eisemana"},
    "ei_planificacion_obra": {"eiplanificacionobra"},
    "ei_fecha_ejecucion": {"eifechaejecucion"},
}
NORMALIZED_ALIASES = {
    _normalize_name(alias): field
    for field, aliases in PROPERTY_ALIASES.items()
    for alias in aliases
}


def _property_values(product) -> dict[str, object]:
    values = {}
    for property_set in get_psets(product, should_inherit=True).values():
        for name, value in property_set.items():
            field = NORMALIZED_ALIASES.get(_normalize_name(name))
            if field and value is not None:
                values[field] = value
    return values


def _as_quantity(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    match = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text)
    return float(match.group()) if match else None


def _as_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _is_executed(planning_status: object, execution_date: object) -> bool:
    status = _normalize_name(planning_status)
    return status in {"ejecutado", "completado", "completed", "true", "1"} or bool(
        _as_text(execution_date)
    )


def load_ifc_model(
    file_path: str,
    include_geometry: bool = True,
    element_ids: set[int] | None = None,
) -> dict:
    """Lee propiedades y, cuando se solicita, mallas de los elementos IFC."""
    model = ifcopenshell.open(str(Path(file_path)))
    products = model.by_type("IfcElement")
    records = []

    for product in products:
        properties = _property_values(product)
        weight = _as_quantity(properties.get("weight"))
        description = _as_text(
            properties.get("assembly_description")
            or getattr(product, "Description", None)
            or getattr(product, "Name", None)
        )
        assembly_code = _as_text(properties.get("assembly_code"))
        execution_date = properties.get("ei_fecha_ejecucion")
        planning_status = properties.get("ei_planificacion_obra")
        records.append({
            "ifc_id": product.id(),
            "global_id": _as_text(getattr(product, "GlobalId", None)),
            "assembly_code": assembly_code,
            "assembly_description": description,
            "nombre": _as_text(getattr(product, "Name", None)),
            "kg": weight,
            "cantidad_total": weight or 0.0,
            "unidad": "kg",
            "ejecutado": _is_executed(planning_status, execution_date),
            "ei_zona": _as_text(properties.get("ei_zona")),
            "ei_subcontratista": _as_text(properties.get("ei_subcontratista")),
            "ei_semana": _as_text(properties.get("ei_semana")),
            "ei_planificacion_obra": _as_text(planning_status),
            "ei_fecha_ejecucion": _as_text(execution_date),
            "fecha": _as_text(execution_date),
        })

    geometries = {}
    geometry_products = (
        products
        if element_ids is None
        else [product for product in products if product.id() in element_ids]
    )
    if include_geometry:
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_WORLD_COORDS, True)
        iterator = ifcopenshell.geom.iterator(
            settings,
            model,
            max(1, min(4, os.cpu_count() or 1)),
            include=geometry_products,
        )
        if iterator.initialize():
            while True:
                shape = iterator.get()
                express_id = int(str(shape.id).split("-", 1)[0])
                geometry = shape.geometry
                geometries[express_id] = {
                    "vertices": tuple(geometry.verts),
                    "faces": tuple(geometry.faces),
                }
                if not iterator.next():
                    break

    return {
        "records": records,
        "geometries": geometries,
        "elements_without_geometry": (
            len(geometry_products) - len(geometries) if include_geometry else None
        ),
    }


def build_ifc_figure(model_data: dict, visible_ids: set[int]):
    """Construye una vista 3D interactiva con colores por estado de ejecución."""
    import plotly.graph_objects as go

    records_by_id = {
        record["ifc_id"]: record for record in model_data["records"]
    }
    vertices_by_status = {True: [], False: []}
    faces_by_status = {True: [], False: []}
    hover_by_status = {True: [], False: []}

    for express_id in visible_ids:
        geometry = model_data["geometries"].get(express_id)
        record = records_by_id.get(express_id)
        if not geometry or not record:
            continue

        status = record["ejecutado"]
        offset = len(vertices_by_status[status]) // 3
        vertices_by_status[status].extend(geometry["vertices"])
        faces_by_status[status].extend(
            tuple(index + offset for index in geometry["faces"][index:index + 3])
            for index in range(0, len(geometry["faces"]), 3)
        )
        hover_text = (
            f"{record['assembly_code'] or 'Sin código'} · "
            f"{record['assembly_description'] or record['nombre']}<br>"
            f"Peso: {record['kg'] or 0.0:,.2f} Kg<br>"
            f"Estado: {'Ejecutado' if status else 'Pendiente'}"
        )
        hover_by_status[status].extend(
            [hover_text] * (len(geometry["vertices"]) // 3)
        )

    figure = go.Figure()
    for status, label, color in (
        (False, "Pendiente / sin ejecutar", "#AAB4C0"),
        (True, "Ejecutado", "#00A0E3"),
    ):
        vertices = vertices_by_status[status]
        faces = faces_by_status[status]
        if not vertices or not faces:
            continue
        coordinates = [vertices[axis::3] for axis in range(3)]
        figure.add_trace(go.Mesh3d(
            x=coordinates[0],
            y=coordinates[1],
            z=coordinates[2],
            i=[face[0] for face in faces],
            j=[face[1] for face in faces],
            k=[face[2] for face in faces],
            color=color,
            name=label,
            hovertext=hover_by_status[status],
            hoverinfo="text",
            flatshading=True,
            showscale=False,
        ))

    figure.update_layout(
        height=650,
        margin={"l": 0, "r": 0, "t": 10, "b": 0},
        scene={
            "aspectmode": "data",
            "xaxis": {"title": "", "visible": False},
            "yaxis": {"title": "", "visible": False},
            "zaxis": {"title": "", "visible": False},
        },
        legend={"orientation": "h", "y": 1.02, "x": 0},
    )
    return figure
