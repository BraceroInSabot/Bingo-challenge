import pandas as pd
import xlsxwriter as xlw
from random import shuffle

# open archive
df = pd.read_excel("Bingo_skyrim.xlsx")
bingo = []

try:
    players_numb = int(input("Player's quantity: "))
except (TypeError, ValueError):
    print("Invalid number! Please insert a valid Integer number.")


def main():
    """
    Returns a list variable with every item on 1º column of the file attached in df variable
    """

    for row in df['Missões']:
        bingo.append(str(row))

    shuffle(bingo)

    for item in bingo:
        if len(bingo) > 25:
            bingo.pop()

    return bingo


def convert():
    """
    Returns a list to Excel (xlsx) convertion with items and players name
    """

    workbook = xlw.Workbook("Desafios.xlsx")
    worksheet = workbook.add_worksheet()

    row, column = 1, 0

    worksheet.write(0, 0, "Challenges")

    for item in bingo:
        worksheet.write(row, column, item)
        row += 1

    def players(p_numb):
        """
        Returns a boolean, it shows if the function where properly finished

            Paramters:

                p_numb: int = Number of players in the game

        """
        p_row = 0
        p_column = 1

        for num in range(0, p_numb):
            try:
                player = str(input("Player's name: "))
            except (TypeError, ValueError):
                print("Invalid type! Please insert a valid player name.")
                return players(players_numb)

            worksheet.write(p_row, p_column, player)
            p_column += 1

        return players is True

    players(players_numb)

    if players:
        workbook.close()
        print("\n\nResult:\n\n------------------------------------------------------------------")
        print(f"{pd.read_excel('Desafios.xlsx')}")
    else:
        print("Something gone wrong! Please, fill the inputs again.")
        players(players_numb)


if __name__ == "__main__":
    main()
    convert()
