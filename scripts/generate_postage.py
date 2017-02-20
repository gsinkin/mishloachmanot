import argparse
import csv
from io import BytesIO
import os
import subprocess
import urllib.request

import easypost

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch


PAGESIZE = (3.5 * inch, 5.5 * inch)
HEIGHT, WIDTH = PAGESIZE


def iterate_csv(csv_path):
    with open(csv_path, 'rU') as infile:
        reader = csv.DictReader(infile)
        for row in reader:
            yield row


def remote_tempfile(path):
    try:
        os.remove(path)
    except:
        pass


def refund_postage(shipments):
    for shipment in shipments:
        try:
            print("Refunding purchased postage: {0}".format(
                shipment.tracking_code))
            shipment.refund()
        except:
            print("Error refunding postage for: {0}".format(shipment))
            continue


def generate_shipments(address, parcel, csv_path):
    shipments = []
    for row in iterate_csv(csv_path):
        print("Creating shipment to: {0}".format(row["SendTo"]))
        shipments.append(easypost.Shipment.create(
            from_address=address,
            parcel=parcel,
            to_address={
                "name": row["SendTo"],
                "street1": row["Address"],
                "street2": row["Address2"],
                "city": row["City"],
                "state": row["State"][:2].upper(),
                "zip": row["Zip"],
                "verify_strict": ["delivery"]
            },
            options={"label_size": "4x6", "label_format": "PDF"}))
    return shipments


def purchase_postage(shipments):
    try:
        for index, shipment in enumerate(shipments):
            print("Purchasing postage to: {0}".format(
                shipment.to_address.name))
            shipment.buy(rate=shipment.lowest_rate(["USPS"], ["Priority"]))
            print("Downloading label {0}".format(
                shipment.postage_label.label_url))
            response = urllib.request.urlopen(
                shipment.postage_label.label_url, timeout=10)
            with open(
                    "labels/ROW_{0}_{1}_LABEL.pdf".format(
                        str(index).zfill(3), shipment.tracking_code),
                    "wb"
            ) as outfile:
                outfile.write(response.read())

    except Exception as error:
        print("Error purchasing postage for {0}: {1}".format(
            shipment.id, error))
        refund_postage(shipments)


def generate_notes(shipments, csv_path):
    for index, row in enumerate(iterate_csv(csv_path)):
        shipment = shipments[index]
        print("Generating note for: {0}".format(shipment.to_address.name))
        pdf_bytes = BytesIO()
        pdf_canvas = canvas.Canvas(pdf_bytes, pagesize=PAGESIZE)
        pdf_canvas.setFont("Helvetica", 10)
        pdf_canvas.drawString(
            1 * inch, HEIGHT - 1 * inch, row["CBI Message"])
        pdf_canvas.drawString(
            1 * inch, HEIGHT - 3 * inch, row["SendingFrom"])
        pdf_canvas.drawString(
            1 * inch, HEIGHT - 4 * inch, row["Generic Message"])
        pdf_canvas.save()
        pdf_bytes.seek(0)
        with open(
                "notes/ROW_{0}_{1}_NOTE.pdf".format(
                    str(index).zfill(3), shipment.tracking_code),
                "wb"
        ) as outfile:
            outfile.write(pdf_bytes.read())


def merge_labels_and_notes(shipments):
    for index, shipment in enumerate(shipments):
        print("Merging label and note for: {0}".format(
            shipment.to_address.name))
        label_path = "labels/ROW_{0}_{1}_LABEL.pdf".format(
            str(index).zfill(3), shipment.tracking_code)
        note_path = "notes/ROW_{0}_{1}_NOTE.pdf".format(
            str(index).zfill(3), shipment.tracking_code)
        label_note_path = "results/ROW_{0}_{1}_LABEL_AND_NOTE.pdf".format(
            str(index).zfill(3), shipment.tracking_code)
        command_args = [
            "pdfjam", "--landscape", "--offset", "'1cm 0cm'",
            "--nup", "2x1", label_path, note_path,
            "--outfile", label_note_path]
        subprocess.check_output(command_args)


def write_results(shipments, csv_path):
    results_path = "results/compiled_results.csv"
    print("Writing CSV results to ".format(results_path))
    with open(results_path, 'w') as outfile:
        for index, row in enumerate(iterate_csv(csv_path)):
            if index == 0:
                fieldnames = sorted(row.keys()) + [
                    "Tracking Code", "Label And Note"]
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)
                writer.writeheader()
            shipment = shipments[index]
            label_note_path = "results/ROW_{0}_{1}_LABEL_AND_NOTE.pdf".format(
                str(index).zfill(3), shipment.tracking_code)
            row.update({"Tracking Code": shipment.tracking_code,
                        "Label And Note": label_note_path})
            writer.writerow(row)


def run(from_address_id, parcel_id, csv_path):
    try:
        from_address = easypost.Address.retrieve(from_address_id)
    except easypost.Error as error:
        raise Exception(
            "Error requesting from_address_id: {0}".format(error))

    try:
        parcel = easypost.Parcel.retrieve(parcel_id)
    except easypost.Error as error:
        raise Exception(
            "Error requesting parcel_id: {0}".format(error))
    if not os.path.exists(csv_path):
        raise Exception("csv-path does not exist")
    shipments = generate_shipments(from_address, parcel, csv_path)
    purchase_postage(shipments)
    try:
        generate_notes(shipments, csv_path)
        merge_labels_and_notes(shipments)
        write_results(shipments, csv_path)
    except Exception as error:
        print("Failure: {0}".format(error))
        refund_postage(shipments)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate USPS postage and note")
    parser.add_argument(
        '--api-key', help='EasyPost API Key', required=True)
    parser.add_argument(
        '--from-address-id', help='Origin EasyPost Address ID (adr_XXXX...)',
        required=True)
    parser.add_argument(
        '--parcel-id', help='EasyPost Parcel ID (pcl_XXXX...)',
        required=True)
    parser.add_argument(
        '--csv-path', help='Addresses and notes CSV', required=True)
    arguments = parser.parse_args()
    easypost.api_key = arguments.api_key
    run(arguments.from_address_id, arguments.parcel_id, arguments.csv_path)
