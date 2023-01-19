#!/usr/bin/env python3
import click
from dotenv import load_dotenv
from emailer.emailer import mail_results
from mysql_tunnel.mysql_tunnel import TunnelSQL
# import pprint
import os
import sys
from titlecase import titlecase
import traceback
from xlrd import open_workbook
import xlwt
from xlutils.copy import copy

xlsfile = ''
dump_opr = False
dump_css = False
verbosity = 0

states = {
    'Alaska': 'AK',
    'Alabama': 'AL',
    'Arkansas': 'AR',
    'American Samoa': 'AS',
    'Arizona': 'AZ',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'District of Columbia': 'DC',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Iowa': 'IA',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Massachusetts': 'MA',
    'Maryland': 'MD',
    'Maine': 'ME',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Missouri': 'MO',
    'Mississippi': 'MS',
    'Montana': 'MT',
    'National': 'NA',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Nebraska': 'NE',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'Nevada': 'NV',
    'New York': 'NY',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Puerto Rico': 'PR',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Virginia': 'VA',
    'Virgin Islands': 'VI',
    'Vermont': 'VT',
    'Washington': 'WA',
    'Wisconsin': 'WI',
    'West Virginia': 'WV',
    'Wyoming': 'WY',
    'Newfoundland and Labrador': 'NL',
    'Prince Edward Island': 'PE',
    'Nova Scotia': 'NS',
    'New Brunswick': 'NB',
    'Quebec': 'QC',
    'Ontario': 'ON',
    'Manitoba': 'MB',
    'Saskatchewan': 'SK',
    'Alberta': 'AB',
    'British Columbia': 'BC',
    'Yukon': 'YT',
    'Northwest Territories': 'NT',
    'Nunavut': 'NU',
}


def dbg(level, text):
    """
    Levels
    0 = no output
    1 = minimal output
    2 = verbose outupt
    3 = very verbose outupt
    4 = show database dumps
    """
    if verbosity > (level - 1):
        print(text)

def resolve_flag(env_var, default):
    """convert enviromntal variable to True False
       return default value if no string"""
    if os.getenv(env_var):
        return [False, True][os.getenv(env_var) != ""]
    return default

def resolve_text(env_var, default):
    """convert enviromntal variable to text string
       return default value if no string"""
    if os.getenv(env_var):
        return os.getenv(env_var)
    return default

def resolve_int(env_var, default):
    return int(resolve_text(env_var, default))

def read_workbook():
    # Read boat/dealer/model from spreadsheet
    # use book.release_resources() before saving
    book = open_workbook(xlsfile, formatting_info=True, on_demand=True)

    # dont assume it is the first sheet
    # scan thru all sheets looking for 'DEALER'
    sx = 0
    for i in range(0, len(book.sheet_names()) - 1):
        if (book.sheet_names()[i] == 'DEALER'):
            sx = i

    # read-only copy  sh.cell(row,col).value sh.cell_value(row,col)
    sh = book.sheet_by_index(sx)

    # to write to file  wb.save('filename')
    wb = copy(book)

    # write-only copy   ws.write(row,col,'value')
    ws = wb.get_sheet(sx)

    # build dictionary of hullserial to row
    hulls = {}
    nulls = 0

    for rx in range(sh.nrows):
        hull = sh.cell_value(rowx=rx, colx=0)

        if (hull[:3] != 'NRB'):
            nulls += 1
            if nulls > 6:
                break
            else:
                continue
        hulls[hull] = rx

    return book, hulls, sh, wb, ws


def fetch_oprs_and_csss(db):
    # select all records from the OPR table
    sql = "SELECT * FROM wp_nrb_opr ORDER BY id"
    oprs = db.execute(sql)

    # select all records from the CS table
    sql = "SELECT * FROM wp_nrb_cs_survey ORDER BY id"
    csss = db.execute(sql)

    if dump_opr:
        print("OPRS [{}]".format(len(oprs)))
        for opr in oprs:
            print(
                "  opr  {:14.14}   {:20.20}   {:22.22} {}   "
                "{:20.20}   {:25.25} {:30.30}".format(
                    opr['hull_serial_number'],
                    opr['dealership'],
                    opr['model'],
                    opr['date_delivered'],
                    opr['first_name'],
                    opr['last_name'],
                    opr['agency']))
    if dump_css:
        if dump_opr:
            print("\n\n\n")
        print("CSS [{}]".format(len(csss)))
        for css in csss:
            print("  css  {:14.14}   {:20.20}   {:22.22}   {}".format(
                css['hull_serial_number'],
                css['dealership'],
                css['model'],
                css['date_purchased'],
            ))

    return oprs, csss


def process_sheet(data, hulls, col, sh, ws):
    """
    For logic see gist:
        https://gist.github.com/northriverboats/bd05796844dee5ecdb493880b5e5e01d
    """
    font_size_style = xlwt.easyxf(
        'font: name Garmond, bold off, height 240;')
    date_font_size_style = xlwt.easyxf(
        'font: name Garmond, bold off, height 240;')
    date_font_size_style.num_format_str = 'mm/dd/yyyy'
    changed = 0

    output = (
        ("", "\n\n\n")[col] +
        "-----------------------------------------------------------"
        "---------------------------------------------------------------------"
        "---------------------------------------------------------------\n"
        "|" + " " * 83 +
        ('O P R \'s  U P D A T E D', 'C S S \'s  U P D A T E D')[col] +
        " " * 83 + '|' +
        "\n-----------------------------------------------------------"
        "---------------------------------------------------------------------"
        "---------------------------------------------------------------\n"
        "| Hull           | Lastname        | Firstname  "
        "| Phone                "
        "| Mailing Address                                    "
        "| Street Address                                     | Purchased  |\n"
        "---------------------------------------------------------------------"
        "---------------------------------------------------------------------"
        "-----------------------------------------------------\n")

    # pp = pprint.PrettyPrinter(indent=4)
    # col 0/4 = opr/css mode
    # opr 0/2 = opr done
    # css 0/1 = css done
    #               0   1  2  3  4  5  6  7
    # truth_table = [T, F, F, F, T, F, T, F]  # CSS priority

    # OPR priority
    truth_table = [True, True,  False, False, True, False, False, False]

    for datum in data:
        rx = hulls.get(
            datum.get('hull_serial_number')[:3] +
            datum.get('hull_serial_number')[4:9] +
            datum.get('hull_serial_number')[10:], 0)
        if (rx):  # rx is row on sheet where datum hull_serial_number shows up
            opr_char = sh.cell_value(rx, 19)
            css_char = sh.cell_value(rx, 20)
            opr_flag = (0, 2)[opr_char != '']
            css_flag = (0, 1)[css_char != '']
            flag = (col * 4) + opr_flag + css_flag
            dbg(3,
                  '  Processing row: {:05d}  hull: {}  flag: {:03b}'.format(
                      rx, datum.get('hull_serial_number'), flag))
            if flag == 6:
                changed += 1
                ws.write(rx, 19 + col, 'X', font_size_style)
            if truth_table[flag]:
                changed += 1
                homephone = str(datum.get('phone_home', '')).upper()
                workphone = str(datum.get('phone_work', '')).upper()
                if (homephone == 'NA' or
                        homephone == 'N/A' or
                        homephone == 'NONE'):
                    homephone = ''
                if (workphone == 'NA' or
                        workphone == 'N/A' or
                        workphone == 'NONE'):
                    workphone = ''

                ws.write(rx,
                         1,
                         titlecase(datum.get('last_name', '')),
                         font_size_style)
                ws.write(rx,
                         2,
                         titlecase(datum.get('first_name', '')),
                         font_size_style)
                ws.write(rx,
                         3,
                         (workphone, homephone)[bool(homephone)],
                         font_size_style)
                ws.write(rx,
                         4,
                         titlecase(datum.get('mailing_address', '')),
                         font_size_style)
                ws.write(rx,
                         5,
                         titlecase(datum.get('mailing_city', '')),
                         font_size_style)
                ws.write(rx,
                         6,
                         states.get(datum.get('mailing_state', ''), ''),
                         font_size_style)
                ws.write(rx,
                         7,
                         datum.get('mailing_zip', '').upper(),
                         font_size_style)
                if col == 0:
                    ws.write(rx,
                             8,
                             titlecase(datum.get('street_address', '')),
                             font_size_style)
                    ws.write(rx,
                             9,
                             titlecase(datum.get('street_city', '')),
                             font_size_style)
                    ws.write(rx,
                             10,
                             states.get(datum.get('street_state', ''), ''),
                             font_size_style)
                    ws.write(rx,
                             11,
                             datum.get('street_zip', '').upper(),
                             font_size_style)
                ws.write(rx,
                         12,
                         datum.get('email_address', datum.get('email', '')),
                         font_size_style)
                ws.write(rx,
                         13,
                         datum.get('date_delivered', ''),
                         date_font_size_style)
                ws.write(rx,
                         19 + col,
                         'X',
                         font_size_style)

                mailing_address = (
                    titlecase(
                        datum.get('mailing_address', '') + ', ' +
                        datum.get('mailing_city', '')) + ', ' +
                    states.get(datum.get('mailing_state', ''), '') + ', ' +
                    datum.get('mailing_zip', '').upper())
                if len(mailing_address) == 6:
                    mailing_address = ''

                street_address = (
                    titlecase(
                        datum.get('street_address', '') + ', ' +
                        datum.get('street_city', '')) + ', ' +
                    states.get(datum.get('street_state', ''), '') + ', ' +
                    datum.get('street_zip', '').upper())
                if len(street_address) == 6:
                    street_address = ''

                if col == 0:
                    date_thing = datum.get('date_delivered', '01/01/01')
                else:
                    date_thing = datum.get('date_purchased', '01/01/01')

                output1 = (
                    "| %-12s | %-15s | %-10s | %-20s |"
                    " %-50s | %-50s | %-10s | %s\n" % (
                        datum.get('hull_serial_number', ''),
                        titlecase(datum.get('last_name', ''))[:15],
                        titlecase(datum.get('first_name', ''))[:10],
                        (workphone, homephone)[bool(homephone)][:20],
                        mailing_address, street_address, date_thing, rx))
                dbg(2, output1.replace('\n', ''))
                output += output1

    return output, changed

@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug verbosity/do not'
              'save verbosity')
@click.option(
    '--verbose', '-v', default=0, type=int, help='verbosity level 0-4')
@click.option(
    '--dumpopr', is_flag=True, help='dump opr table')
@click.option(
    '--dumpcss', is_flag=True, help='dump css table')
def main(debug, verbose, dumpopr, dumpcss):
    global xlsfile
    global verbosity
    global dump_opr
    global dump_css

    # set python environment
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
    else:
        # we are running in a normal Python environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    # load environmental variables
    load_dotenv(bundle_dir + '/.env')

    if os.getenv('HELP'):
      with click.get_current_context() as ctx:
        click.echo(ctx.get_help())
        ctx.exit()

    verbosity = resolve_int('VERBOSE', verbose)
    no_save = resolve_flag('DEBUG', debug)
    dump_opr = resolve_flag('DUMPOPR', dumpopr)
    dump_css = resolve_flag('DUMPCSS', dumpcss)

    xlsfile = os.getenv('XLSFILE')

    if verbosity > 0:
        try:
            print(f"{xlsfile} is {os.path.getsize(xlsfile)} bytes in size")
        except OSError as e:
            print(f"{xlsfile} is not found")

    try:
        silent = verbosity < 3
        db = TunnelSQL(silent, cursor='DictCursor')
        oprs, csss = fetch_oprs_and_csss(db)
        book, hulls, sh, wb, ws = read_workbook()
        output_2, changed_2 = process_sheet(csss, hulls, 1, sh, ws)
        dbg(3, "CSS's changed: {}\n".format(changed_2))
        output_1, changed_1 = process_sheet(oprs, hulls, 0, sh, ws)
        dbg(3, "OPR's changed: {}\n".format(changed_1))
        output = output_1 + output_2
        changed = changed_1 + changed_2

        if (changed and not no_save):
            wb.save(xlsfile)
            mail_results(
                'OPR to Warranty Spreadsheet Update',
                '<pre>' + output + '</pre>')
            dbg(1, output)
        else:
            dbg(1, 'No changes File Not Saved')
    except OSError:
        mail_results(
            'OPR to Warranty Spreadsheet is open',
            'OPR to Warranty Spreadsheet is open, '
            'spreadsheet can not be updated'
            "<br /><br /><pre>" + traceback.format_exc() + "</pre>")
    except Exception as e:
        mail_results(
            'OPR to Warranty Spreadsheet Processing Error',
            '<p>Spreadsheet can not be updated due to script error:<br />\n' +
            str(e) + '</p>'
            "<br /><br /><pre>" + traceback.format_exc() + "</pre>")
    finally:
        db.close()
    sys.exit(0)

if __name__ == "__main__":
    main()
