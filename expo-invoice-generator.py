import csv
from fpdf import FPDF, XPos, YPos, Align
from fpdf.enums import MethodReturnValue, WrapMode
from collections import defaultdict

FONT_SZ_HEADER = 14
FONT_SZ_FOOTER = 7

FONT_SZ_TABLE_HEADER = 10
FONT_SZ_TABLE_ROW = 8

FONT_SZ_NORMAL = 10

FONT_PATH = '/usr/share/fonts/truetype/dejavu'

ROW_HEIGHT_HEADER = 8
ROW_HEIGHT_ROW = 6

def format_price(p):
    return f"{p:,.2f} €".translate(str.maketrans(",.", ".,"))

def format_booking_number(bn_int):
    bn = str(bn_int)
    return f"EX-{bn[0:3]}-{bn[3:7]}"

# Initialize FPDF class
class PDF(FPDF):
    widths = (30, 68, 60, 25)
    local_page_no = 0

    def new_document(self):
        self.local_page_no = self.page_no()
    
    def header(self):
        self.set_font("DejaVuSans", "B", FONT_SZ_HEADER)
        self.cell(0, 10, "Relevé des visites organisées par le MNHN", align=Align.C, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(8)

    def footer(self):
        # Go to 1.5 cm from bottom
        self.set_y(-15)
        self.set_font("DejaVuSans", "", FONT_SZ_FOOTER)
        self.cell(0, 10, f"Page {self.page_no() - self.local_page_no + 1}", align=Align.R, new_x=XPos.RIGHT, new_y=YPos.TOP)

    def row(self, height, cells, border="B",
            weights=("", "", "", ""),
            aligns=("L", "L", "L", "L")):

        # Measure cell height in lines
        cell_col_lines = []
        for i, el in enumerate(zip(cells, self.widths, aligns, weights)):
            text, width, a, weight = el
            if i < len(cells)-1:
                new_x = XPos.RIGHT
                new_y = YPos.TOP
            else: # last col
                new_x = XPos.LMARGIN
                new_y = YPos.NEXT

            self.set_font("DejaVuSans", weight, FONT_SZ_TABLE_ROW)
            cell_col_lines.append(
                self.multi_cell(width, height, text, border=border, align=a,
                                new_x=new_x, new_y=new_y, wrapmode=WrapMode.CHAR,
                                dry_run=True, output=MethodReturnValue.LINES))

        max_lines = max([len(cell_lines) for cell_lines in cell_col_lines])

        # Actually output cells
        for i, el in enumerate(zip(cell_col_lines, self.widths, aligns, weights)):
            cell_lines, width, a, weight = el
            if i < len(cells)-1:
                new_x = XPos.RIGHT
                new_y = YPos.TOP
            else: # last col
                new_x = XPos.LMARGIN
                new_y = YPos.NEXT

            self.set_font("DejaVuSans", weight, FONT_SZ_TABLE_ROW)
            empty_lines = max_lines - len(cell_lines)
            text = "".join(cell_lines + ["\n " * empty_lines])

            self.multi_cell(width, height, text, border=border, align=a,
                            new_x=new_x, new_y=new_y, wrapmode=WrapMode.CHAR)

    def add_commune_data(self, commune, events):
        self.set_font("DejaVuSans", "B", FONT_SZ_TABLE_HEADER)
        self.row(ROW_HEIGHT_HEADER,
                 ["Date / # Activité", "Responsable", "Nom", "Prix"],
                 weights=["I", "I", "I", "I"],
                 aligns=["L", "L", "L", "R"])

        total_sum = 0

        self.set_font("DejaVuSans", "", FONT_SZ_TABLE_ROW)
        for entry in events:
            total_sum += entry['price']

            self.row(ROW_HEIGHT_ROW, [
                format_booking_number(entry['booking_number']),
                entry['responsable'],
                entry['activity'],
                format_price(entry['price'])
            ], border="", weights=["", "B", "B", ""], aligns=["L", "L", "L", "R"])
            self.row(ROW_HEIGHT_ROW,
                     [entry['datetime'], entry['customer_name'], "", ""],
                     weights=["", "I", "", ""])

        # Total sum for the commune
        self.set_font("DejaVuSans", "B", FONT_SZ_TABLE_ROW)
        self.row(ROW_HEIGHT_ROW, ["", "", "Total", format_price(total_sum)],
                 weights=["", "", "B", "B"],
                 aligns=["L", "L", "L", "R"])

def generate_reports(csv_file_path):
    data = defaultdict(list)

    with open(csv_file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            activity = row["Offer\nName"]
            customer_invoice_address_name = row["Customer\nInvoice address name"]
            data[customer_invoice_address_name].append({
                'cia_name': row["Customer\nInvoice address name"],
                'cia_street': row["Customer\nInvoice address street"],
                'cia_zip': row["Customer\nInvoice address postal code"],
                'cia_city': row["Customer\nInvoice address city"],
                'activity': activity,
                'responsable': row["Booker\nFull name"],
                'customer_name': row["Customer\nName"],
                'booking_number': row["Booking\nNumber"],
                'datetime': row["Offer\nEnd date & time"],
                'price': float(row["Reservation\nPrice"])
            })

    pdf = PDF()
    
    pdf.set_margins(16, 20, 16) # L T R 
    for customer_invoice_address_name, activities in data.items():
        
        pdf.add_font('DejaVuSans', '', f'{FONT_PATH}/DejaVuSans.ttf')
        pdf.add_font('DejaVuSans', 'B', f'{FONT_PATH}/DejaVuSans-Bold.ttf')
        pdf.add_font('DejaVuSans', 'I', f'{FONT_PATH}/DejaVuSans-Oblique.ttf')
        pdf.add_page()
        pdf.new_document()

        # Initial content
        pdf.set_font("DejaVuSans", "B", FONT_SZ_NORMAL)
        pdf.multi_cell(0, 5, f"Adresse de facturation pour la commune {customer_invoice_address_name} et les visites ci-dessous:", 0, 'L')
        pdf.ln(2)

        pdf.set_font("DejaVuSans", "", FONT_SZ_NORMAL)
        pdf.set_x(pdf.get_x() + 10)

        a0 = activities[0]
        pdf.multi_cell(0, 5, f"{a0['cia_name']}\n{a0['cia_street']}\n{a0['cia_zip']} {a0['cia_city']}", 0, 'L')
        pdf.ln(10)


        # Adding activity data
        pdf.add_commune_data(customer_invoice_address_name, activities)

    filename_suffix = customer_invoice_address_name.replace(' ', '-').replace("'", "")
    pdf.output(f"facture_{filename_suffix}.pdf")
    print(f"Generated PDF for {customer_invoice_address_name}")

if __name__ == '__main__':
    import sys
    generate_reports(sys.argv[1])
