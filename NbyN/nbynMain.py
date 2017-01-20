# Javier Vazquez


# from NbyN.nbynClass import nbyn
import nbynClass

def main():
    n = input("Table dimension: ")
    nTable = nbynClass.nbyn(int(n))
    print(nTable.sizeOfTable())
    nTable.addingHeader()
    nTable.populatingTable()


if __name__ == '__main__':
    main()