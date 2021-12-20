import openpyxl

class Card:
    """
    A class to model a trading card. Card object stores name, type, HP, shininess,
    moves and associated damage.

    """
    def __init__(self, theName, theType, theHP, isShiny, theMoves):
        """
        Constructs attributes for Card object.

        Parameters
        ----------
            theName: string (default: "")
                Name of card
            theType: string (default: "")
                Card type
            theHP: int (default: 0)
                Card HP
            isShiny: bool (default: None)
                True if card is shiny; False if not
            theMoves: dict (default: {})
                Dictionary where keys are move names and values are damage inflicted
                by move

        """
        self.name = theName
        self.hp = theHP
        self.moves = theMoves
        self.shiny = isShiny
        self.type = theType

    def __str__(self):
        return f"""Name:           {self.name}
Type:           {self.type}
HP:             {self.hp}
Moves\Damage:   {" | ".join(f"{k} {v}" for k, v in self.moves.items())}
Shiny:          {"Yes" if self.shiny else "No"}
"""

    def getCardAverageDamage(self):
        """
        Returns average damage over all the card's moves.

        Parameters - None
        Returns - float

        """
        cardTotalDamage = 0
        moveCount = 0
        for damage in self.moves.values():
            cardTotalDamage += damage
            moveCount += 1
        cardAverageDamage = cardTotalDamage / moveCount

        return cardAverageDamage

    def getName(self):
        """
        Returns the name of the card.

        Parameters - None
        Returns - string

        """
        return self.name

    def isShiny(self):
        """
        Returns True if card is shiny, returns False if not.

        Parameters - None
        Returns - bool

        """
        return self.shiny

    def getType(self):
        """
        Returns the type of the card.

        Parameters - None
        Returns - string

        """
        return self.type

    def getHp(self):
        """
        Returns HP of card.

        Parameters - None
        Returns - int

        """
        return self.hp

    def getMoves(self):
        """
        Returns dictionary where keys are moves and values are damage values.

        Parameters - None
        Returns - dict

        """
        return self.moves

class Deck:
    """
    A class modelling a deck of cards. Stores a list of card objects. Implements
    methods for adding and removing cards, inputting cards from a formatted *.xlsx,
    and saving deck as *.xlsx. Also implements methods for viewing: whole deck,
    deck by type or shiny status, and most powerful card in deck.

    """
    def __init__(self):
        """
        Constructs attributes for Deck object.

        Parameters - None

        """
        self.cards = []

    def addCard(self, theCard):
       """
       Adds card to list of cards in Deck object.

       Parameters - theCard (Card)
       Returns - None

       """
       self.cards.append(theCard)
       print("Card added to deck.")

    def rmCard(self, theCard):
        """
        Removes card from list of cards in Deck object, if present.

        Parameters - theCard (Card)
        Returns - None

        """
        try:
            self.cards.remove(theCard)
            print("Card removed.")
        except ValueError:
            print("Card not present in deck.")

    def inputFromFile(self, fileName):
        """
        Takes card parameters from formatted *.xlsx file, generates card objects
        and inserts them into deck.

        Parameters - FileName (string)
        Returns - None

        """

        print("Inputting from file...\n")

        # Open *.xlsx using openpxyl and access sheet
        try:
            book = openpyxl.load_workbook(fileName)
            sheet = book.active
        except FileNotFoundError:
            print("File not found.")
            return
        except Exception as exp:
            print("File failed to open due to unspecified reason.", exp)
            return

        # Parse rows from sheet into list containing list of values, from cell objects, for each card:
        try:
            rowsAsValues = []
            targetRow = 0
            for row in sheet.rows:
                # For each row in sheet representing a card, create a new list within rowsAsValues
                rowsAsValues.append([])
                # Add values from row into list within rowsAsValues
                for cell in row:
                    rowsAsValues[targetRow].append(cell.value)
                # Move to next row
                targetRow += 1

        except Exception:
            print("Failed to get values from spreadsheet. Please check spreasheet formatting matches "
                  "sampleDeck.xlsx")

        try:
            # Create card object by indexing to appropriate locations in rowsAsValues
            for rowOfValues in rowsAsValues[1:]:  # Remove title row

                # Test for empty rows + ignore
                if rowOfValues[0] == None:
                    continue

                else:
                    # Get values
                    cardName = rowOfValues[0]   # name
                    cardType = rowOfValues[1]   # type
                    cardHp = rowOfValues[2]     # hp

                    # Test if card is shiny and create variable
                    cardShiny = False
                    if rowOfValues[3] == 1:
                        cardShiny = True

                    # Create move:dmg dict:
                    moveDamageDict = {}
                    # Define initial positions in each row in rowOfValues corresponding
                    # to first Move:Damage pair
                    pos1 = 4
                    pos2 = 5
                    # Iterate over Move:Damage pairs
                    for i in range(5):
                        # Test for empty cells
                        if rowOfValues[pos1] == None or rowOfValues[pos2] == None:
                            pass
                        else:
                            # Append Move:Damage to moveDamageDict
                            moveDamageDict[rowOfValues[pos1]] = rowOfValues[pos2]
                            # Increment positions to move to next move:damage pair
                            pos1 += 2
                            pos2 += 2

                    # Generate card object from above variables
                    card = Card(cardName, cardType, cardHp, cardShiny, moveDamageDict)

                    # Add card to self.cards
                    self.addCard(card)

        except Exception:
            print(f"Failed to generate a card object {card.getName()}. Check spreadsheet formatting matches "
                  "sampleDeck.xlsx")

        else:
            print("\nInputting from file complete.\n")

    def getAverageDamage(self):
        """
        Returns the average damage of the average damage of all cards in the deck, rounded to 1 decimal place.

        Parameters - None
        Returns - float

        """

        sumAverage = 0
        cardCount = 0

        for card in self.cards:
            cardAverageDamage = card.getCardAverageDamage()
            sumAverage += cardAverageDamage
            cardCount += 1

        deckAverageDamage = sumAverage / cardCount

        return round(deckAverageDamage, 1)

    def getMostPowerful(self):
        """
        Returns the card with the highest average damage in the deck.

        Parameters - None
        Returns - Card

        """

        # Create cardName:avgDmg dict
        cardDamageDict = {}

        # Populate dict
        for card in self.cards:
            cardName = card.getName()
            cardAverageDamage = card.getCardAverageDamage()
            cardDamageDict[cardName] = cardAverageDamage

        # Find largest avgDmg and get cardName from dict
        mostPowerfulName = max(cardDamageDict, key= lambda x: cardDamageDict[x])

        # Use cardName to return most powerful card
        for card in self.cards:
            if card.getName() == mostPowerfulName:
                return card

    def viewAllCards(self):
        """
        Prints all cards in deck.

        Parameters - None
        Returns - None

        """
        print("******************************************")
        print("CARDS IN DECK:")
        print("******************************************")
        for card in self.cards:
            print(card)

    def viewAllShinyCards(self):
        """
        Prints all cards in deck which are shiny.

        Parameters - None
        Returns - None

        """

        print("******************************************")
        print("SHINY CARDS IN DECK:")
        print("******************************************")
        for card in self.cards:
            if card.isShiny():
                print(card)

    def viewAllByType(self, theType):
        """
        Prints all cards in deck which are of the specified type.

        Parameters - theType (string)
        Returns - None

        """

        print("******************************************")
        print(f"CARDS OF TYPE {theType.upper()} IN DECK:")
        print("******************************************")
        for card in self.cards:
            if card.getType().lower() == theType.lower():
                print(card)

    def getCards(self):
        """
        Returns all cards in the deck as a collection.

        Parameters - None
        Returns - list

        """
        return self.cards

    def saveToFile(self, fileName):
        """
        Saves deck as *.xlsx file.

        Parameters - fileName (string)
        Returns - None

        """
        try:
            # Generate title row as a TUPLE for input to xlsx
            titleRow = ("Name", "Type", "HP", "Shiny", "Move Name 1", "Damage 1",
                        "Move Name 2", "Damage 2", "Move Name 3", "Damage 3",
                        "Move Name 4", "Damage 4", "Move Name 5", "Damage 5")

            # Initialize list, which will itself contain lists of card values
            cardRows = []


            for card in self.cards:

                # Get simple values (Name, Type, Hp... *but not moves*) for each card
                cardName = card.getName()
                cardType = card.getType()
                cardHp = card.getHp()

                # Convert True for Shiny to 1 or False to 0
                cardShiny = 0
                if card.isShiny():
                    cardShiny = 1

                # Get list of all move:damage pairs in moves dict and append each
                # to a list like [move, damage, move, damage,...]
                cardMovesList = []
                for pair in card.moves.items():
                    for item in pair:
                        cardMovesList.append(item)

                # Make list of simple columns and extend with move/damage columns
                allCardAttributes = [cardName, cardType, cardHp, cardShiny]
                allCardAttributes.extend(cardMovesList)

                # Convert list of card values to tuple and append to cardRows
                cardRows.append(tuple(allCardAttributes))

            # Open spreadsheet
            book = openpyxl.Workbook()
            sheet = book.active

            # Write title and card rows to spreadsheet
            sheet.append(titleRow)

            for card in cardRows:
                sheet.append(card)

            # Save new sheet
            book.save(fileName)

        except Exception:
            print("Failed to save to file.")

        else:
            # Print "saved" message
            print(f"Deck saved to file {fileName}.")
