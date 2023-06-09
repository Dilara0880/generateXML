from xml.dom import minidom
import pandas as pd
from lxml import etree
import logging
import requests
from bs4 import BeautifulSoup


def read_xlsx(input_file):
    """
        This function takes a name of input Excel file,
        reads it and puts data to DataFrame df and file_name.
    """
    logging.info('Loading Excel data...')
    df = pd.read_excel(input_file, header=4)
    df_header = pd.read_excel(input_file, header=None, nrows=3)
    file_name = df_header.at[2, 1]
    return df, file_name


def get_exchange_rate(date):
    """
        This function takes a date as input and
        retrieves the exchange rate of USD on that date
        from the Central Bank of Russia's website.
    """
    date = date.strftime('%d.%m.%Y')
    url = f'https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={date}'
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    rate = soup.find('td', string='USD').find_next_sibling('td').find_next_sibling('td').find_next_sibling('td').text
    rate = rate.replace(',', '.')
    return float(rate)


def create_xml(df, file_name):
    """
        This function takes a pandas DataFrame as input and
        creates an XML file based on the data in the DataFrame.
    """
    logging.info('Generating XML data...')
    root = etree.Element("CERTDATA")
    filename = etree.SubElement(root, "FILENAME")
    filename.text = file_name
    envelope = etree.SubElement(root, "ENVELOPE")

    # iterator for each row
    for _, row in df.iterrows():
        ecert = etree.SubElement(envelope, "ECERT")
        etree.SubElement(ecert, "CERTNO").text = row['Ref no']
        etree.SubElement(ecert, "CERTDATE").text = row['Issuance Date'].strftime('%Y-%m-%d')
        etree.SubElement(ecert, "STATUS").text = row['Status']
        etree.SubElement(ecert, "IEC").text = str(row['IE Code'])
        etree.SubElement(ecert, "EXPNAME").text = f"\"{row['Client']}\""
        etree.SubElement(ecert, "BILLID").text = row['Bill Ref no']
        etree.SubElement(ecert, "SDATE").text = row['SB Date'].strftime('%Y-%m-%d')
        etree.SubElement(ecert, "SCC").text = row['SB Currency']
        if row['SB Amount'] % 1 == 0:
            etree.SubElement(ecert, "SVALUE").text = f"{int(row['SB Amount'])}"
        else:
            etree.SubElement(ecert, "SVALUE").text = f"{row['SB Amount']:.2f}"

        # getting the USD rate for SB Date
        exchange_rate = get_exchange_rate(pd.to_datetime(row['SB Date']))
        svalue_usd = round(float(row['SB Amount']) / exchange_rate, 2)
        etree.SubElement(ecert, "SVALUEUSD").text = str(svalue_usd)

    xml_str = etree.tostring(root, encoding='UTF-8')
    xml_parsed = minidom.parseString(xml_str)
    xml_formatted = xml_parsed.toprettyxml(indent="\t")
    xml_formatted = xml_formatted.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')
    logging.info('Writing data to XML file...')
    with open("output.xml", "w", encoding="utf-8") as xml_file:
        xml_file.write(xml_formatted)
    logging.info('Created XML file "output.xml"')


def main():
    """
        The main function of the program.
    """

    # enable logging with time
    logging.basicConfig(level=logging.INFO, filename='app.log', filemode='a',
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %H:%M:%S')

    df, file_name = read_xlsx('test_input.xlsx')
    create_xml(df, file_name)


if __name__ == "__main__":
    main()
