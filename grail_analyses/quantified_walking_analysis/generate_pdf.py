import io

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak

from .config import movement_names


def generate_pdf(data, save_path):

    doc = SimpleDocTemplate(
        save_path, pagesize=letter, leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.4 * inch, bottomMargin=0.5 * inch)
    elements = []

    identification = [["Numéro de dossier"], ["Nom", ""], ["Prénom", ""], ["Date", ""]]
    identification_table = Table(identification, colWidths=[106, 400])
    identification_table.setStyle(
        TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
    )
    elements.append(identification_table)
    elements.append(Spacer(1, 20))  # Add an empty spacer for separation

    _create_measure_table(elements, data=data, title="Cinématique", level="kinematics")
    _create_measure_table(elements, data=data, title="Moment", level="moment")
    _create_measure_table(elements, data=data, title="Puissance", level="power")

    doc.build(elements)
    _create_page_numbered_pdf(save_path)


def _create_measure_table(elements, data, title: str, level: str):
    if level not in ("kinematics", "moment", "power"):
        raise ValueError("Wrong level")

    title_table = Table([[title]], colWidths=[475])
    title_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ]))
    elements.append(title_table)

    phases = [["", "Phase d'Appui", "Phase d'Oscillations"]]
    phases_table = Table(phases, colWidths=[140, 183, 183])
    phases_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('LINEBEFORE', (0, 0), (0, -1), 2, colors.black),
        ('LINEAFTER', (-1, 0), (-1, -1), 2, colors.black),
        ('LINEAFTER', (0, 0), (-1, -1), 1, colors.black),
        ('LINEABOVE', (0, 0), (-1, 0), 2, colors.black),
    ]))
    elements.append(phases_table)

    measures = [["", "Min (deg)", "Max (deg)", "Range (deg)", "Min (deg)", "Max (deg)", "Range (deg)"]]
    for name, plane in zip(("Sagittal", "Frontal", "Transversal"), ("sagittal", "frontal", "transversal")):
        _populate_measures(measures, data=data, name=name, level=level, plane=plane)
    measure_table = Table(measures, colWidths=[140, 61, 61, 61, 61, 61, 61])
    n_sagittal = len(movement_names[level]["sagittal"]) - 1
    n_frontal = len(movement_names[level]["frontal"]) - 1
    measure_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, n_sagittal + 3), (0, n_sagittal + 3), 'Helvetica-Bold'),
        ('FONTNAME', (0, n_sagittal + n_frontal + 5), (0, n_sagittal + n_frontal + 5), 'Helvetica-Bold'),
        ('LINEBEFORE', (0, 0), (0, -1), 2, colors.black),
        ('LINEAFTER', (-1, 0), (-1, -1), 2, colors.black),
        ('LINEAFTER', (-4, 0), (-4, -1), 1, colors.black),
        ('LINEAFTER', (0, 0), (0, -1), 1, colors.black),
        ('LINEBELOW', (0, n_sagittal + 2), (-1, n_sagittal + 2), 1, colors.black),
        ('LINEBELOW', (0, n_sagittal + n_frontal + 4), (-1, n_sagittal + n_frontal + 4), 1, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('LINEBELOW', (0, -1), (-1, -1), 2, colors.black),
    ]))
    elements.append(measure_table)

    elements.append(PageBreak())


def _populate_measures(measures: list, data: dict, name: str, level: str, plane: str):
    if level not in ("kinematics", "moment", "power"):
        raise ValueError("Wrong level")
    if plane not in ("sagittal", "frontal", "transversal"):
        raise ValueError("Wrong plane")

    measures.append([name])
    names = movement_names[level][plane]
    for i in range(len(names)):
        row = [
            names[i]["fr"],
            round(data[names[i]["en"]]["Minimum Appui"], 2),
            round(data[names[i]["en"]]["Maximum Appui"], 2),
            round(data[names[i]["en"]]["Range Appui"], 2),
            round(data[names[i]["en"]]["Minimum Oscillation"], 2),
            round(data[names[i]["en"]]["Maximum Oscillation"], 2),
            round(data[names[i]["en"]]["Range Oscillation"], 2),
        ]
        measures.append(row)


def _create_page_numbered_pdf(file_path):
    existing_pdf = PdfFileReader(open(file_path, "rb"))
    num_pages = existing_pdf.getNumPages()
    output = PdfFileWriter()
    for i in range(num_pages):
        page = existing_pdf.getPage(i)

        # Create a canvas and add page number
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        can.drawString(520, 20, f"{i + 1} of {num_pages}")
        can.drawString(50, 20, "Signature:____________________")

        can.save()

        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        page.mergePage(new_pdf.getPage(0))

        # Add the page to the output pdf
        output.addPage(page)

    # Write the output pdf
    with open(file_path, "wb") as output_stream:
        output.write(output_stream)
