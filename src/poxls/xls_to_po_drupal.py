from collections import OrderedDict
import os
import sys
import time
import re
import click
import polib
import openpyxl
from . import ColumnHeaders
import pycountry


def save(output_file, catalog):
    """Save catalog to a PO file (Python 3)."""
    output_file.write(str(catalog))
    output_file.write("\n")  # add final newline


def po_timestamp(filename):
    local = time.localtime(os.stat(filename).st_mtime)
    offset = -(time.altzone if local.tm_isdst else time.timezone)
    return '%s%s%s' % (
        time.strftime('%Y-%m-%d %H:%M', local),
        '-' if offset < 0 else '+',
        time.strftime('%H%M', time.gmtime(abs(offset)))
    )


# === Helpers nettoyage / normalisation ===
def _clean_text(s):
    """Cleans anything that can break detection: Excel leading apostrophe, NBSP, zero-width, trim."""
    if s is None:
        return None
    s = str(s)
    if s.startswith("'"):           # apostrophe réelle
        s = s[1:]
    s = s.replace("\u00A0", " ")    # NBSP
    s = s.replace("\u200B", "")     # zero-width space
    return s.strip()

def _norm_count(s):
    """Ensures we always have '@count ' with a space."""
    s = _clean_text(s)
    if not s:
        return s
    if "@count" in s and "@count " not in s:
        s = s.replace("@count", "@count ")
    s = re.sub(r"\s{2,}", " ", s)
    return s


@click.command()
@click.argument('locale', required=True)
@click.argument('input_file',
        type=click.Path(exists=True, readable=True),
        required=True)
@click.argument('output_file', type=click.File('w', encoding='utf-8'), required=True)
@click.argument('project', required=True)
@click.argument('version', required=True)
@click.option("--debug", is_flag=True, help="Affiche les décisions de couplage @count")
def main(locale, input_file, output_file, project, version, debug):
    """
    Convert a XLS(X) file to a .PO file
    """
    language = pycountry.languages.get(alpha_2=locale).name
    click.echo(f'Processing {language} translations...')
    book = openpyxl.load_workbook(input_file)
    catalog = polib.POFile()
    catalog.header = 'Futurium translation for Drupal by CNECT.R3'
    catalog.metata_is_fuzzy = True
    catalog.metadata = OrderedDict()
    catalog.metadata['Project-Id-Version'] = f'{project} ({version})'
    catalog.metadata['POT-Creation-Date'] = po_timestamp(input_file)
    catalog.metadata['PO-Revision-Date'] = 'YYYY-mm-DD HH:MM+ZZZZ'
    catalog.metadata['Content-Type'] = 'text/plain; charset=utf-8'
    catalog.metadata['Content-Transfer-Encoding'] = '8bit'
    catalog.metadata['Language-Team'] = language
    catalog.metadata['MIME-Version'] = '1.0'

    # Plural-Forms
    if locale in ['bg', 'da', 'de', 'el', 'en', 'es', 'et', 'fi', 'hu', 'it', 'nl', 'pt', 'sv']:
        plural_form = 'nplurals=2; plural=(n!=1);'
    elif locale in ['fr']:
        plural_form = 'nplurals=2; plural=(n>1);'
    elif locale in ['cs', 'sk']:
        plural_form = 'nplurals=3; plural=((n==1)?(0):(((n>=2)&&(n<=4))?(1):2));'
    elif locale in ['lv']:
        plural_form = 'nplurals=3; plural=(n%10==1 && n%100!=11 ? 0 : n != 0 ? 1 : 2);'
    elif locale in ['ro']:
        plural_form = 'nplurals=3; plural=(n==1 ? 0 : (n==0 || (n%100 > 0 && n%100 < 20)) ? 1 : 2);'
    elif locale in ['lt']:
        plural_form = 'nplurals=3; plural=(n%10==1 && n%100!=11 ? 0 : n%10>=2 && (n%100<10 || n%100>=20) ? 1 : 2);'
    elif locale in ['pl']:
        plural_form = 'nplurals=3; plural=(n==1 ? 0 : n%10>=2 && n%10<=4 && (n%100<10 || n%100>=20) ? 1 : 2);'
    elif locale in ['hr']:
        plural_form = 'nplurals=3; plural=(n%10==1 && n%100!=11 ? 0 : n%10>=2 && n%10<=4 && (n%100<10 || n%100>=20) ? 1 : 2);'
    elif locale in ['sl']:
        plural_form = 'nplurals=4; plural=(n%100==1 ? 0 : n%100==2 ? 1 : n%100==3 || n%100==4 ? 2 : 3);'
    elif locale in ['mt']:
        plural_form = 'nplurals=4; plural=(n==1 ? 0 : n==0 || ( n%100>1 && n%100<11) ? 1 : (n%100>10 && n%100<20 ) ? 2 : 3);'
    elif locale in ['ga']:
        plural_form = 'nplurals=5; plural=n==1 ? 0 : n==2 ? 1 : (n>2 && n<7) ? 2 :(n>6 && n<11) ? 3 : 4;'
    else:
        print(f"Unknown plural form for {locale}")
        sys.exit(1)
    catalog.metadata['Plural-Forms'] = plural_form

    for sheet in book.worksheets:
        if sheet.max_row < 2:
            continue
        row_iterator = sheet.iter_rows()
        headers = [c.value for c in next(row_iterator)]
        headers = dict((b, a) for (a, b) in enumerate(headers))
        msgctxt_column = headers.get(ColumnHeaders.msgctxt)
        msgid_column = headers.get(ColumnHeaders.msgid)
        tcomment_column = headers.get(ColumnHeaders.tcomment)
        comment_column = headers.get(ColumnHeaders.comment)
        msgstr_column = headers.get(locale)
        if msgid_column is None or msgstr_column is None:
            continue

        rows = list(row_iterator)
        n = len(rows)
        i = 0
        while i < n:
            row_vals = [_clean_text(c.value) for c in rows[i]]
            msgid = _norm_count(row_vals[msgid_column])
            if not msgid:
                i += 1
                continue

            # --- Detect an @count pair ---
            if "@count" in msgid and (i + 1) < n:
                next_vals = [_clean_text(c.value) for c in rows[i + 1]]
                next_msgid = _norm_count(next_vals[msgid_column])

                if debug:
                    click.echo(f"[DBG] i={i}: msgid={repr(msgid)} | next={repr(next_msgid)}")

                if next_msgid and "@count" in next_msgid:
                    entry = polib.POEntry(msgid=msgid, msgid_plural=next_msgid)
                    entry.msgstr_plural[0] = _norm_count(row_vals[msgstr_column])
                    raw_plural_tr = _norm_count(next_vals[msgstr_column])

                    if raw_plural_tr and "/" in raw_plural_tr:
                        plural1, plural2 = raw_plural_tr.split("/", 1)
                        entry.msgstr_plural[1] = _norm_count(plural1)
                        entry.msgstr_plural[2] = _norm_count(f"@count {plural2.strip()}")
                    elif raw_plural_tr and "(" in raw_plural_tr:
                        plural1 = _norm_count(raw_plural_tr.split("(")[0])
                        parts = raw_plural_tr.split()
                        if len(parts) > 1:
                            plural2 = "".join([ch for ch in parts[1] if ch.isalpha()])
                            entry.msgstr_plural[1] = plural1
                            entry.msgstr_plural[2] = _norm_count(f"@count {plural2}")
                        else:
                            entry.msgstr_plural[1] = plural1
                    else:
                        entry.msgstr_plural[1] = raw_plural_tr

                    if msgctxt_column is not None and row_vals[msgctxt_column]:
                        entry.msgctxt = row_vals[msgctxt_column]
                    if tcomment_column and row_vals[tcomment_column]:
                        entry.tcomment = row_vals[tcomment_column]
                    if comment_column and row_vals[comment_column]:
                        entry.comment = row_vals[comment_column]

                    catalog.append(entry)

                    if debug:
                        click.echo(f"[DBG]  → COMBINED: {repr(entry.msgid)} / {repr(entry.msgid_plural)}")
                        click.echo(f"[DBG]     msgstr[0]={repr(entry.msgstr_plural.get(0))} | msgstr[1]={repr(entry.msgstr_plural.get(1))}")

                    i += 2
                    continue

            # --- Simple case ---
            entry = polib.POEntry(msgid=msgid)
            entry.msgstr = _norm_count(row_vals[msgstr_column])
            catalog.append(entry)

            if debug:
                click.echo(f"[DBG] i={i}: SIMPLE -> {repr(entry.msgid)}")

            i += 1

    if not catalog:
        click.echo('No messages found, aborting', err=True)
        sys.exit(1)

    save(output_file, catalog)


if __name__ == '__main__':
    main()
