"""Generate a synthetic, explicitly delimited text-layer mailroom PDF; no client data."""
import argparse
import hashlib
import json
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--output', type=Path, required=True)
    args = ap.parse_args()
    if args.output.exists():
        raise FileExistsError(f'Refusing to overwrite source: {args.output.name}')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    font_dir = Path(__file__).resolve().parents[1] / 'assets/fonts'
    regular = TTFont('MailArial', str(font_dir / 'DejaVuSans.ttf'))
    bold = TTFont('MailArialBold', str(font_dir / 'DejaVuSans-Bold.ttf'))
    pdfmetrics.registerFont(regular)
    pdfmetrics.registerFont(bold)
    names = ['Nguyễn Thị Hà', 'José García', 'Zoë O’Connor', 'François Noël', 'Chloé Martin',
             'Müller & Söhne', 'André Silva', 'Renée Laurent', 'Søren Jensen', 'Ana María Díaz',
             'Bùi Minh Đức', 'Nguyễn Thị Hà', 'Alice / Bob', 'A_B % Partner', 'Taylor Lee']
    statuses = ['Received', 'Pending Review', 'Approved', 'On Hold', 'Closed']
    lengths = [1, 2, 3, 1, 2, 1, 1, 2, 1, 2, 1, 3, 1, 1, 2]
    docs = [dict(Document_Type=f'Type_{i % 11 + 1:02}', Person_Name=name,
                 Status=statuses[i % len(statuses)], Document_ID=f'MR-2026-{i+1:04}',
                 Date=f'2026-11-{i+1:02}', Page_Count=lengths[i]) for i, name in enumerate(names)]
    required = ''.join(names) + ''.join(statuses)
    if any(ord(char) not in regular.face.charToGlyph or ord(char) not in bold.face.charToGlyph for char in required):
        raise ValueError('Packaged fonts lack a required Unicode glyph; refusing corrupt sample output')
    c = canvas.Canvas(str(args.output), pagesize=(612, 792), invariant=1)
    c.setTitle('Synthetic Mailroom Sample - 15 documents / 24 pages')
    c.setAuthor('Local mailroom proof of concept')
    for doc in docs:
        for local_page in range(1, doc['Page_Count'] + 1):
            c.setFillColor(colors.HexColor('#183153'))
            c.rect(42, 716, 528, 34, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont('MailArialBold', 13)
            c.drawString(54, 728, 'MAILROOM | SYNTHETIC TRAINING SAMPLE')
            c.setFillColor(colors.HexColor('#17283C'))
            c.setFont('MailArialBold', 15)
            c.drawString(48, 685, 'Correspondence record' if local_page == 1 else 'Correspondence continuation')
            fields = [('Document Type', doc['Document_Type']), ('Document ID', doc['Document_ID'])]
            if local_page == 1:
                fields += [('Person Name', doc['Person_Name']), ('Status', doc['Status']), ('Date', doc['Date'])]
            c.setFont('MailArial', 11)
            y = 653
            for label, value in fields:
                c.drawString(48, y, f'{label}: {value}')
                y -= 23
            c.setFont('MailArialBold', 11)
            c.drawString(48, y - 13, 'Processing details' if local_page == 1 else 'Supporting schedule')
            c.setFont('MailArial', 10)
            c.drawString(48, y - 39, 'This sample tests document boundaries, Unicode fields and local routing.')
            c.drawString(48, y - 56, 'The identifiers and people shown here are synthetic demonstration records.')
            data = [['Line', 'Description', 'Units', 'Amount (USD)'],
                    ['01', f'Handling task {doc["Document_ID"]}', str(local_page), '24.50'],
                    ['02', 'Verification and registration', '1', '10.00'],
                    ['03', 'Archive reference / attachment', '2', '0.00']]
            t = Table(data, colWidths=[42, 306, 55, 113], rowHeights=27)
            t.setStyle(TableStyle([('FONTNAME', (0, 0), (-1, -1), 'MailArial'),
                                  ('FONTNAME', (0, 0), (-1, 0), 'MailArialBold'),
                                  ('FONTSIZE', (0, 0), (-1, -1), 9),
                                  ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E6ECF3')),
                                  ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#17283C')),
                                  ('LINEBELOW', (0, 0), (-1, 0), 0.7, colors.HexColor('#8292A6')),
                                  ('LINEBELOW', (0, 1), (-1, -1), 0.3, colors.HexColor('#D8E1EB')),
                                  ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
                                  ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            t.wrapOn(c, 516, 200)
            t.drawOn(c, 48, y - 200)
            c.setFont('MailArial', 9)
            c.setFillColor(colors.HexColor('#5C6B7B'))
            c.drawString(48, 74, 'Synthetic text-layer PDF | No OCR performed | Not customer correspondence')
            c.setFont('MailArialBold', 10)
            c.drawRightString(564, 52, f'Page {local_page} of {doc["Page_Count"]}')
            c.showPage()
    c.save()
    print(json.dumps({'synthetic': True, 'documents': len(docs), 'pages': sum(lengths),
                      'output': args.output.name, 'sha256': hashlib.sha256(args.output.read_bytes()).hexdigest()}))


if __name__ == '__main__':
    main()
