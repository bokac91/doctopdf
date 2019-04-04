import comtypes.client
import time
import shutil
import math
import os
from os import listdir
from os.path import isfile, join


def pdf_exists(path, file_name):
    dir_files = [fd for fd in listdir(path) if isfile(join(path, fd))]
    for f_i in dir_files:
        if f_i == file_name:
            return True


# Putanja glavnog direktorijuma
dir_path = os.path.dirname(os.path.realpath(__file__))

# TODO: Napraviti propisno ispravnu funkciju za ignorisanje/brisanje? __MACOSX fajlova
# Kratko ciscenje od _MACOSX fajlova
for f in os.listdir(dir_path):
    if f == "__MACOSX":
        shutil.rmtree(f)

# VISUAL STUFF
print("Welcome! Program is starting.")
num_of_files = 0
# Prilikom prebrojavanja fajlova za obradu, potrebno je IGNORISATI backup fajlove!
for root, dirs, files in os.walk(dir_path):
    for file in files:
        if file.endswith(".doc") or file.endswith(".docx"):
            if file.startswith("~$"):
                continue
            else:
                num_of_files += 1

print("Traversing trough \'" + dir_path + "\'")
time.sleep(.3)

# SLUCAJ 1: ako ne postoje fajlovi za konverziju, kraj programa!
if num_of_files == 0:
    print("There're no potential files to be converted. Exiting...")
    exit()

# Pomeraj za procenat
percentage_divider = 100/num_of_files
total_progress = 0
print("Amount of (potential) files for conversion: " + str(num_of_files))
print("------------------------------------------------------")
time.sleep(.3)
print("Progress: " + str(total_progress) + "%")

# Pocetak merenja vremena programa - na ovom mestu, zbog sleep-va iznad
start = time.time()

# Kreiranja objekta za konverzije
word = comtypes.client.CreateObject('Word.Application')
# word.Visible = True
# time.sleep(2)

# PDF format, brojac slucajeva i pomocne za segmente loadinga
wdFormatPDF = 17
i = 0
l_c = 0
x = 5

# Fajl sa punim izvestajem o odradjenom poslu
rf = open("report.txt", "w+", encoding="utf-8")

# SLUCAJ 2: glavna petlja obrade fajlova
for root, dirs, files in os.walk(dir_path):
    for file in files:

        # BUG: Neizbezna situacija - backup, nevidljivi fajlovi ~$
        if file.startswith('~$'):
            continue

        elif file.endswith(".doc"):
            i += 1
            rf.write("Found " + str(i) + ": [" + file + "]\n")

            # Apsolutna putanja ulaznog (Word) fajla
            p_in = os.path.join(root, file)
            in_file = os.path.abspath(p_in)

            # Apsolutna putanja izlaznog (PDF) fajla
            new_name = file.replace(".doc", r".pdf")
            out_file = os.path.join(root, new_name)

            # SIGURNA promena progresa - bez obzira na ishod!
            total_progress += percentage_divider

            # Provera postojanja .pdf fajla
            if pdf_exists(root, new_name):
                rf.write("SKIPPED -> .pdf already exists!\n")
                rf.write("------------------------------------\n")
                if l_c % x == 0:
                    print("Current: %.1f" % total_progress + "%")
                continue

            # Konverzija
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()

            if math.ceil(total_progress) != 100:
                if l_c % x == 0:
                    print("Current: %.1f" % total_progress + "%")

            rf.write("Success! File created: \"" + new_name + "\"\n")
            rf.write("------------------------------------\n")

        elif file.endswith(".docx"):
            i += 1
            rf.write("Found " + str(i) + ": [" + file + "]\n")

            p_in = os.path.join(root, file)
            in_file = os.path.abspath(p_in)

            new_name = file.replace(".docx", r".pdf")
            out_file = os.path.join(root, new_name)

            # SIGURNA promena progresa - bez obzira na ishod!
            total_progress += percentage_divider

            if pdf_exists(root, new_name):
                rf.write("SKIPPED -> .pdf already exists!\n")
                rf.write("------------------------------------\n")
                if l_c % x == 0:
                    print("Current: %.1f" % total_progress + "%")
                continue

            docx = word.Documents.Open(in_file)
            docx.SaveAs(out_file, FileFormat=wdFormatPDF)
            docx.Close()
            if math.ceil(total_progress) != 100:
                if l_c % x == 0:
                    print("Current: %.1f" % total_progress + "%")

            rf.write("Success! File created: \"" + new_name + "\"\n")
            rf.write("------------------------------------\n")
    l_c += 1

# word.Visible = False
word.Quit()

end = time.time()
total = end - start

print("Progress: 100%")
print("------------------------------------------------------")
print("Amount of files parsed: %i/%i" % (i, num_of_files))
print("Total execution runtime: %.2f" % total + "s.")
if i == num_of_files:
    print("Conversion fully complete!")

rf.write("Total execution runtime: %.2f" % total + "s.")
rf.close()
