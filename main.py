from xml.dom import minidom
import pandas as pd
from lxml import etree
import logging


def read_xlsx(input_file):
    logging.info('Loading Excel data...')
    df = pd.read_excel(input_file, header=4)
    df_header = pd.read_excel(input_file, header=None, nrows=3)
    file_name = df_header.at[2, 1]
    return df, file_name


def create_xml(df, file_name):
    logging.info('Generating XML data...')
    root = etree.Element("CERTDATA")
    filename = etree.SubElement(root, "FILENAME")
    filename.text = file_name
    envelope = etree.SubElement(root, "ENVELOPE")

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

    xml_str = etree.tostring(root, encoding='UTF-8')
    xml_parsed = minidom.parseString(xml_str)
    xml_formatted = xml_parsed.toprettyxml(indent="\t")
    xml_formatted = xml_formatted.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')
    logging.info('Writing data to XML file...')
    with open("output.xml", "w", encoding="utf-8") as xml_file:
        xml_file.write(xml_formatted)
    logging.info('Created XML file "output.xml"')


def main():
    logging.basicConfig(level=logging.INFO, filename='app.log', filemode='w',
                        format='%(name)s - %(levelname)s - %(message)s')
    df, file_name = read_xlsx('test_input.xlsx')
    create_xml(df, file_name)


if __name__ == "__main__":
    main()
