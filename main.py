import PyPDF2
import re
import unicodecsv as csv

pdfFileObj = open('input.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
data = []
numPages = pdfReader.getNumPages()

print("Loaded PDF File. Pages Found: " + str(numPages))

for i in range(0, numPages):
    print("Processing Page " + str(i))
    pageObj = pdfReader.getPage(i)
    pageText = pageObj.extractText()

    pageSplit = pageText.split("\n \n \n \n")

    speaker = re.split("\n \n[0-9]*\n \n", pageSplit[0])
    pageSplit[0] = speaker[1]

    for item in pageSplit:
        if len(item) > 10:
            itemText = item

            if "\n \n \n" in item:
                itemSplit = item.split("\n \n \n")
            elif "\n \nPhD\n \n" in item:
                itemSplit = item.split("\n \nPhD\n \n")
            elif "\n \nJD\n \n" in item:
                itemSplit = item.split("\n \nJD\n \n")

            name = itemSplit[0].replace("\n", "").replace("**", "")

            if " \n \n" in itemSplit[1]:
                itemFields = itemSplit[1].split(" \n \n")
            else:
                itemFields = itemSplit[1].split("\n \n")

            position = itemFields[0].replace("\n", "")
            itemFields.pop(0)

            address = ""
            email = ""
            telephone = ""

            for field in itemFields:
                if not ("Email" or "Telephone") in field:
                    address += field + " "
                elif ("Email" or "Telephone") in field:
                    itemContact = field.split("\n \n")
                    for contact in itemContact:
                        if "Email" in contact:
                            email = contact.replace("Email: ", "").replace("\n", "")
                        elif "Telephone" in contact:
                            telephone = contact.replace("Telephone: ", "").replace("\n", "")

            address = address.replace("\n", "")
            data.append([name, position, address, telephone, email])

print("Saving Data to CSV File.")
with open('output.csv', mode='wb') as output:
    csvOutput = csv.writer(output, dialect='excel', encoding='utf_8_sig')
    csvOutput.writerows(data)
print("Task Completed.")


