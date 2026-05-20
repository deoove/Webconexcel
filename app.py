from flask import Flask, render_template, request, redirect, url_for, flash
import openpyxl
import os

app = Flask(__name__)
app.secret_key = "clave_secreta"

EXCEL_FILE = "datos.xlsx"


def get_workbook():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Personas"
        ws.append(["ID", "Nombre", "Apellido", "DNI"])
        wb.save(EXCEL_FILE)
    return openpyxl.load_workbook(EXCEL_FILE)


def save_workbook(wb):
    wb.save(EXCEL_FILE)


def get_next_id(ws):
    ids = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0] is not None]
    return max(ids, default=0) + 1


def get_all_personas():
    wb = get_workbook()
    ws = wb.active
    personas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            personas.append({"id": row[0], "nombre": row[1], "apellido": row[2], "dni": row[3]})
    return personas


def find_row(ws, persona_id):
    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if row[0].value == persona_id:
            return idx
    return None


@app.route("/")
def index():
    personas = get_all_personas()
    return render_template("index.html", personas=personas)


@app.route("/alta", methods=["GET", "POST"])
def alta():
    if request.method == "POST":
        nombre = request.form["nombre"].strip()
        apellido = request.form["apellido"].strip()
        dni = request.form["dni"].strip()
        if not nombre or not apellido or not dni:
            flash("Todos los campos son obligatorios.", "error")
            return render_template("form.html", accion="Alta", persona=None)
        wb = get_workbook()
        ws = wb.active
        new_id = get_next_id(ws)
        ws.append([new_id, nombre, apellido, dni])
        save_workbook(wb)
        flash("Persona dada de alta correctamente.", "success")
        return redirect(url_for("index"))
    return render_template("form.html", accion="Alta", persona=None)


@app.route("/modificar/<int:persona_id>", methods=["GET", "POST"])
def modificar(persona_id):
    wb = get_workbook()
    ws = wb.active
    row_idx = find_row(ws, persona_id)
    if row_idx is None:
        flash("Persona no encontrada.", "error")
        return redirect(url_for("index"))
    if request.method == "POST":
        nombre = request.form["nombre"].strip()
        apellido = request.form["apellido"].strip()
        dni = request.form["dni"].strip()
        if not nombre or not apellido or not dni:
            flash("Todos los campos son obligatorios.", "error")
        else:
            ws.cell(row=row_idx, column=2).value = nombre
            ws.cell(row=row_idx, column=3).value = apellido
            ws.cell(row=row_idx, column=4).value = dni
            save_workbook(wb)
            flash("Persona modificada correctamente.", "success")
            return redirect(url_for("index"))
    persona = {
        "id": ws.cell(row=row_idx, column=1).value,
        "nombre": ws.cell(row=row_idx, column=2).value,
        "apellido": ws.cell(row=row_idx, column=3).value,
        "dni": ws.cell(row=row_idx, column=4).value,
    }
    return render_template("form.html", accion="Modificar", persona=persona)


@app.route("/eliminar/<int:persona_id>", methods=["POST"])
def eliminar(persona_id):
    wb = get_workbook()
    ws = wb.active
    row_idx = find_row(ws, persona_id)
    if row_idx is None:
        flash("Persona no encontrada.", "error")
    else:
        ws.delete_rows(row_idx)
        save_workbook(wb)
        flash("Persona eliminada correctamente.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
