from botcity.web import WebBot, Browser
from botcity.web.bot import By
from tkinter import *
from tkinter import messagebox
from openpyxl import *
import os

class Bot(WebBot):

    # Excel Config
    wb = load_workbook('Results.xlsx')
    ggp = wb['Great Games Playlist']
    gta = wb['Games to Avoid']

    def action(self, execution=None):
        # Config
        self.headless = True
        print("Headless True")
        self.driver_path = 'C:/Users/xsh4d/Documents/Pythons/gameRatingBot/chromedriver.exe'
        self.graph()

    def search(self, gamename):
        # Get game name
        game = gamename

        # Opens the Metacritic website.
        self.browse("https://www.metacritic.com/")

        # Find Search Box
        elem = self.find_element("primary_search_box", By.ID)
        elem.click()

        # Find the Game
        self.paste(game)
        self.enter(500)

        # Filter 4 Games
        self.execute_javascript('document.getElementsByClassName("filter_button game")[0].click();')
        self.wait(500)

    def graph(self):
        root = Tk()
        root.title("Game Rating Bot")
        root.iconbitmap('icon.ico')

        labelOne = Label(text="Welcome to the Game Rating Bot! You give me a name and I give you a number!")
        labelOne.pack()
        labelTwo = Label(text="Please, input the name of the game in the field below:")
        labelTwo.pack()
        userInput = Entry(root, bg="light gray", borderwidth=4, width=40)
        userInput.pack()
        def getInput():
            messagebox.showinfo("GRB 1.0", "Ok, I'm going to fetch its Metacritic rating! \nPlease, wait a moment...")
            self.process(userInput.get())
        startButton = Button(root, text="Search!", padx=10, pady=10, command=getInput)
        startButton.pack()
        def openexcel():
            os.system('start EXCEL.EXE Results.xlsx')
        excelButton = Button(root, text="Open Playlist", padx=10, pady=10, command=openexcel)
        excelButton.pack()

        root.mainloop()

    def addtoexcel(self, ws, name, score):
        ws.append([name, score])
        self.wb.save('Results.xlsx')

    def process(self, gamename):
        try:
            game = gamename
            self.search(game)

            # Fetch the Score
            index = 0
            game_name = None
            while str(game_name).lower().replace(':', '').replace('-', '') not in game.lower().replace(':', '').replace('-', '') or game.lower().replace(':', '').replace('-', '') not in str(game_name).lower().replace(':', '').replace('-', ''):
                game_name = self.execute_javascript(f'return document.getElementsByClassName("search_results module")[0].getElementsByTagName("li")[{index}].getElementsByTagName("a")[0].innerText;')
                index = index + 1

            # Grade / Show
            score = self.execute_javascript(f'return document.getElementsByClassName("search_results module")[0].getElementsByTagName("li")[{index - 1}].getElementsByTagName("span")[0].innerText;')
            if str(score) == "tbd":
                messagebox.showinfo("GRB 1.0", "Hey, I found your game! Its score is yet to be published... Please, try again another time!")
                addtoGTA = 0
                addtoGGP = 0
            elif int(score) <= 30:
                messagebox.showinfo("GRB 1.0", "Hey, I found your game! Its score is: " + score +"!\nIt sucks, sorry.")
                addtoGTA = 1
                messagebox.showinfo("GRB 1.0","I'm going to add it to the Games to Avoid Sheet, you deserve better...")
                addtoGGP = 0
            elif int(score) > 30 and int(score) <= 60:
                messagebox.showinfo("GRB 1.0", "Hey, I found your game! Its score is: " + score +"!\nIt seems it isn't one of best, but I guess it is worth a shot.")
                addtoGTA = messagebox.askyesno("GRB 1.0","Do you want to add it to the Games to Avoid sheet?")
                addtoGGP = messagebox.askyesno("GRB 1.0", "Do you want to add it to the Great Games Playlist sheet?")
            elif int(score) > 60 and int(score) <= 90:
                messagebox.showinfo("GRB 1.0", "Hey, I found your game! Its score is: " + score + "!\nIt definitely goes to the playlist.")
                addtoGGP = messagebox.askyesno("GRB 1.0","Do you want to add it to the Great Games Playlist sheet?")
                addtoGTA = 0
            elif int(score) > 90:
                messagebox.showinfo("GRB 1.0", "Hey, I found your game! Its score is: " + score + "!\nYou should buy it RIGHT NOW! This game is just... PERFECT!")
                messagebox.showinfo("GRB 1.0","I'm adding it to the Great Games Playlist, you better play it!")
                addtoGGP = 1

            # Add to Excel File
            if addtoGGP == 1:
                self.addtoexcel(self.ggp, game, score)
            elif addtoGTA == 1:
                self.addtoexcel(self.gta, game, score)
            elif addtoGTA == 0 and addtoGGP == 0:
                messagebox.showinfo("GRB 1.0","Thanks for using our services!")

            return 1
        except Exception as e:
            messagebox.showerror("GRB 1.0", "Sorry, I ran into some problems... Maybe the game is not in Metacritic's database or is misspelled. In case of a spelling error, please try again with the correct spelling!")
            return 0

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    Bot.main()














