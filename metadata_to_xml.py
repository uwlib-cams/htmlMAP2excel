#!python
"""
This program transforms HTML metadata into Excel. It is extremely
specifically tailored to the HTML in contentdm_maps.

Example HTML: https://uwlib-mig.github.io/contentdm_maps/html/becker.html
"""
import argparse
import logging
import os
import re

# HTML processing library:
from bs4 import BeautifulSoup
# XLSX writing library:
import xlsxwriter


# Logging + options to add more logging if you set `DEBUG` environment
# variable
if os.environ.get('DEBUG', False):
    logging.basicConfig(level=logging.DEBUG)
else:
    logging.basicConfig(level=logging.INFO)

logger = logging.getLogger()


def iterate_table_with_header_row(table):
    """
    Given a BeautifulSoup table tag, read the table. The first row
    is assumed to be the header (hence `_with_header_row`). That row
    is set to be the "key" for each subsequent row.

    Yields a dictionary per row of the table (after skipping the first
    row).
    """
    table_rows = table.find_all('tr')

    # Get the first row + remove it from table_rows (pop)
    header_row = table_rows.pop(0)

    header_row_text = list(map(
        # Get the text only:
        lambda x: x.text.strip(),
        # From any <th> or <td> entity in the row:
        header_row.find_all(['th', 'td'])))

    # For all the remaining rows.
    for row in table_rows:
        row_text = map(
            # Get the text only:
            lambda x: x.text.strip(),
            # From any <th> or <td>
            row.find_all(['th', 'td']))

        # keys - header text
        # values - row text
        row_val = dict(zip(header_row_text, row_text))
        yield row_val


WHITESPACE_RE = re.compile('\s+')
def remove_whitespace(ws_text):
    return WHITESPACE_RE.sub(' ', ws_text)


def convert_html_to_text(soup):
    text = ''

    for element in soup.children:
        if element.name == 'li':
            text += "• " + convert_html_to_text(element) + "\n"
        elif element.name == 'a' and element.has_attr('href'):
            href = element['href']
            text += convert_html_to_text(element) + f" ({href})"
        elif element.name:
            text += convert_html_to_text(element)
        else:
            text += remove_whitespace(element.string)

    return text

def read_table_with_header_col(table):
    """
    Returns a dictionary where the keys are the first columns and
    the values are the second columns, given a BeautifulSoup table.

    Any rows with only one element are considered to be the title of
    the table and get the special key `TABLE_TITLE`.
    """
    table_rows = table.find_all('tr')

    # The resulting dictionary
    values = {}
    
    for row in table_rows:
        row_text = list(map(
            convert_html_to_text,

            # From any <th> or <td> element:
            row.find_all(['th', 'td'])))

        # TODO this is kind of a hack; assuming there's only one row
        # that has only one value, based on what I've seen in the HTML
        if len(row_text) == 1:
            values['TABLE_TITLE'] = row_text[0]
            continue

        assert len(row_text) == 2, "Looking for exactly two columns"
        values[row_text[0]] = row_text[1]

    return values


class HTMLMetadata:
    """
    Object to read/store/access the metadata for a given HTML
    file.
    """
    def __init__(self, html_fh):
        """
        Given a file handle, read the HTML from that file handle
        into this object.
        """
        self.fh = html_fh
        self.header_data = {}
        self.prop_data = {}
        self.field_vals = {}
        self._read_fh()

    def _read_header(self, soup):
        """
        Get the 'all_list_table' table and read it.
        """

        # get the header table
        header_tables = soup.find_all('table', class_='all_list_table')
        assert len(header_tables) == 1, "Need one header table"
        header_table = header_tables[0]

        # read header info
        for row in iterate_table_with_header_row(header_table):
            assert 'Field label' in row, "'Field label' key needs to exist"
            row_key = row['Field label']

            if row_key in self.header_data:
                logger.error("Duplicate header key '%s'; skipping",
                             row_key)
                continue

            self.header_data[row_key] = row

    def _read_prop_tables(self, soup):
        """
        Find the tables with 'prop_table' class and read them.
        """
        # read property table
        for prop_table in soup.find_all('table', class_='prop_table'):
            prop_data = read_table_with_header_col(prop_table)

            # From the special row with only one column, look for the
            # text to the right of the colon. That's the property
            # name.
            table_title = ' '.join(prop_data['TABLE_TITLE'].split())

            (prop_type, prop_name) = table_title.split(':')
            prop_data['PROPERTY_TYPE'] = prop_type.strip()
            prop_data['PROPERTY_NAME'] = prop_name.strip()

            if table_title in self.prop_data:
                logger.error("Duplicate property table '%s'; skipping",
                             table_title)

            self.prop_data[table_title] = prop_data

    def _read_field_configurations(self, soup):
        """
        Get the field_configuration <h3>'s and read the associated
        lists.
        """
        # read field configuration deets
        for field_config in soup.find_all('h3', class_='prop_table_head'):
            # The <ul> is unfortunately not related via a div or
            # anything, so just get the next sibling.
            bullets = field_config.find_next_sibling(string=None)

            assert bullets.name == 'ul', "Need the next element after the <h3> to be a <ul>"

            # Read the text of each list item into 'text_list'
            text_list = list(map(
                lambda x: x.text.strip(),
                bullets.find_all('li')))

            # Now go through the text and split it on ':'. Store it as
            # key/value pairs into field_vals
            field_vals = {}
            for item in text_list:
                (key, val) = list(map(lambda x: x.strip(), item.split(':')))
                field_vals[key] = val

            assert 'Field label' in field_vals, "'Field label' field needs to exist"
            label = field_vals['Field label']

            if label in self.field_vals:
                logger.error("Duplicate field configuration '%s'; skipping",
                             label)
                continue

            self.field_vals[label] = field_vals


    def _read_fh(self):
        soup = BeautifulSoup(self.fh, 'html.parser')

        self._read_header(soup)
        self._read_prop_tables(soup)
        self._read_field_configurations(soup)

    def get_header(self, field_name):
        return self.header_data[field_name]

    def get_field(self, field_name):
        return self.field_vals[field_name]

    def get_fields(self):
        return self.field_vals.values()

    def get_properties(self):
        return self.prop_data.values()


def read_html_metadata(html_file):
    with open(html_file) as html_fh:
        metadata = HTMLMetadata(html_fh=html_fh)
    return metadata


def write_xlsx_row(row_num, col_vals, worksheet):
    """
    Convenience function to write out a whole row to an Excel worksheet.
    """
    for (col_num, col_val) in enumerate(col_vals):
        worksheet.write(
            row_num,
            col_num,
            col_val)


def write_field_worksheet(field_worksheet, metadata):
    """
    Given an XlsxWriter worksheet and metadata object, populate
    the worksheet with the ContentDM fields.
    """
    # for each CONTENTdm field:
    row_num = 0
    write_xlsx_row(
        row_num=row_num,
        col_vals=(
            'Field Label',
            'CDMS DC Map',
	    'CDMS Show Large Field',
	    'CDMS Searchable',
	    'CDMS Hidden',
	    'CDMS Required',
	    'CDMS Controlled Vocabulary'),
        worksheet=field_worksheet)
    
    for field_dict in metadata.get_fields():
        row_num += 1
        col_vals = [field_dict[key]
                    for key in
                    (
                        "Field label",
                        "CONTENTdm setting 'DC map'",
                        "CONTENTdm setting 'Show large field'",
                        "CONTENTdm setting 'Searchable'",
                        "CONTENTdm setting 'Hidden'",
                        "CONTENTdm setting 'Required'",
                        "CONTENTdm setting 'Controlled vocabulary'"
                        )]
        write_xlsx_row(
            row_num=row_num,
            col_vals=col_vals,
            worksheet=field_worksheet)
        

def write_property_data(property_worksheet, metadata):
    """
    Given an XlsxWriter worksheet and a metadata object, populate the worksheet based on the metadata properties.
    """
    row_num = 0
    write_xlsx_row(
        row_num=row_num,
        col_vals=(
            'Property Type',
            'Field Name',
	    'DC Property',
	    'Field Order',
	    'Definition',
	    'Instructions',
	    'Examples',
	    'Show large field',
	    'Searchable',
	    'Hidden',
	    'Required',
	    'Controlled vocabulary'),
        worksheet=property_worksheet)

    for prop_dict in metadata.get_properties():
        row_num += 1

        field_name = prop_dict['PROPERTY_NAME']
        field = metadata.get_field(field_name)
        header = metadata.get_header(field_name)
        col_vals = (
            prop_dict['PROPERTY_TYPE'],
            field_name,
            field["CONTENTdm setting 'DC map'"],
            header['Field order'],
            prop_dict['Property definition'],
            prop_dict['Recording values'],
            prop_dict['Examples / well-formed values'],
            field["CONTENTdm setting 'Show large field'"],
            field["CONTENTdm setting 'Searchable'"],
            field["CONTENTdm setting 'Hidden'"],
            field["CONTENTdm setting 'Required'"],
            field["CONTENTdm setting 'Controlled vocabulary'"],
            )

        write_xlsx_row(
            row_num=row_num,
            col_vals=col_vals,
            worksheet=property_worksheet)

    



def write_metadata_to_xlsx(
        metadata:HTMLMetadata,
        xlsx_file:str):

    if os.path.exists(xlsx_file):
        os.unlink(xlsx_file)

    workbook = xlsxwriter.Workbook(xlsx_file)
    field_worksheet = workbook.add_worksheet('Field Configuration Details')
    write_field_worksheet(
        field_worksheet=field_worksheet,
        metadata=metadata)

    property_worksheet = workbook.add_worksheet('Property data')
    # Worksheet #2, to record property data:
    write_property_data(
        property_worksheet=property_worksheet,
        metadata=metadata)

    workbook.close()



def convert_html_metadata_to_xlsx(
        html_file:str,
        xlsx_file:str):

    metadata = read_html_metadata(html_file)

    write_metadata_to_xlsx(
        metadata=metadata,
        xlsx_file=xlsx_file)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="")

    parser.add_argument('--inhtml',
                        required=True)
    parser.add_argument('--outxlsx',
                        required=True)

    args = parser.parse_args()

    convert_html_metadata_to_xlsx(
        html_file=args.inhtml,
        xlsx_file=args.outxlsx)
