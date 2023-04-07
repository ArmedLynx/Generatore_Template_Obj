from MyClass.Data import Data
from MyClass.Data import CsvData
from MyClass.Template import Template

from os import mkdir
from os.path import exists as path_exists
from os.path import splitext as splitext

import sys, getopt

def ApriData(path):
    # Recupero l'estensione del file data
    file_name, file_extension = splitext(path)
    if file_extension == ".csv" or file_extension == ".CSV":
        #Creo un oggetto contenente il file csv
        return CsvData(path), "csv"
    elif file_extension == ".xlsx" or file_extension == ".XLSX":
        # Creo un oggetto contenente il folgio di lavoro xlsx
        return Data(path), "xlsx"
    else:
        print("# ERR: Formato file data non supportato")
        raise SystemExit

def main(argv):
    opts, args = getopt.getopt(argv, "t:d:o:FAh") # Eseguo il parsing della lista di argomenti per ottenere una lista di tuple (opzione, argomento)

    aiuto = '''
        -t <Path>\tPercorso per il file template
        -d <Path>\tPercorso per il file data
        -o <Path>\tPercorso per la cartella di output
        -F\tForza la riscrittura dei file di configurazione già creati
        -A\tConcatena l'output nello stesso file. Ignora \"-F\" se presente
        -h\tAiuto
        '''
    # Assegno alle variabili il valore di default
    template_file = "./template.txt"
    data_file = "./data.csv"
    out_path = "./outputs"
    overwrite = False
    append = False

    # Assegno alle variabili il valore passato da cmd
    for opt, arg in opts:
        if opt == "-t":
            template_file = arg
        elif opt == "-d":
            data_file = arg
        elif opt == "-o":
            out_path = arg
        elif opt == "-F":
            overwrite = True
        elif opt == "-A":
            append =True
        elif opt == "-h":
            print(aiuto)
            return
        else:
            print("Parametri non validi.\n"+aiuto)

    # Se la cartella di output non esiste la creo
    if path_exists(out_path) == False:
        print(": Creo la cartella di output\n|")
        mkdir(out_path)
    
    # Verifico se il formato del file Data è supportato
    print(": Verifico il formato del file data")
    xls, Formato = ApriData(data_file)
    print(":- Ho aperto il file "+Formato+"\n|")
    
    # Eseguo un ciclo sulla lista contente i filename
    for file in xls.GetFileNames():
        # if append == True:
        #     print("!!!devo concatenare!!!")
        # Se il flag overwrite è False e il file esiste non faccio nulla altrimenti creo il file configurazione
        if overwrite == append == False and path_exists(out_path+"/"+file+".txt") == True:
            print(": Il file "+file+".txt è già presente")
        else:
            # Creo un oggetto contenente il template
            conf = Template(template_file)
            print(":- Apro il template per "+file)
            # Eseguo un ciclo sulla lista contenete i tags
            for tag in xls.GetTags():
                # Sostituisco il tag con il contenuto della cella alle coordinate corrispondenti al file e al tag
                conf.Replace(tag, xls.GetValue(xls.GetCell(file, tag)))
                print(":-- Sostituisco "+tag+" con "+str(xls.GetValue(xls.GetCell(file, tag))))
            # Salvo il file di configurazione
            if append:
                conf.Append(file+".txt", out_path)
            else:
                conf.Save(file+".txt", out_path)
            print(":- Salvo il file di cofigurazione "+file+"\n|")
    
    print(": Ho terminato la preparazione delle configurazioni")
    
# def test():
#     xls = CsvData(r'C:\Users\lspreafico.MATICMINDIT\Desktop\Cartel1.csv')
#     print(xls.GetFileNames())
#     for i in range (1, 3):
#         print(xls.GetFileName(i))
#     print(xls.GetTags())
#     for i in range (1, 4):
#         print(xls.GetTag(i))
#     print(xls.GetValue("B2"))
#     print(xls.GetValue("C3"))
#     print(xls.GetCell("pippo", "<NET>"))
#     print(xls.GetCell("pluto", "<HOST>"))
#     print(xls.GetValue(xls.GetCell("pippo", "<NET>")))
#     print(xls.GetValue(xls.GetCell("pluto", "<HOST>")))

if __name__ == "__main__":
    main(sys.argv[1:])
    # test()