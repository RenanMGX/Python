import status_online
import os
#install necessary libraries
os.system("pip install -U tk")
os.system("apt-get install python3-tk")

#Add the following imports
import tkinter as tk
from tkinter import ttk

#Rest of the code is same as before
#Create the GUI window, listbox, button and functions etc

#To save the list of computers and their statuses to a file, you can use the with open statement and the write() method.
#Add the following function
def save_to_file():
    with open("computer_status.txt", "w") as file:
        for i in range(len(computers)):
            file.write(computers[i] + ": " + statuses[i] + "\n")

#Add a button to the GUI to call the save_to_file function
save_button = ttk.Button(root, text="Save to File", command=save_to_file)
save_button.pack()

#To split the script into multiple parts and execute them separately, you can use the if __name__ == "__main__": statement.
#Move the code that creates the GUI window, listbox, button and functions etc inside the if statement
if 1 == 1:
# GUI code here
    root = tk.Tk()
    root.title("Computer Status")
    listbox = tk.Listbox(root)
    listbox.pack(fill=tk.BOTH, expand=True)
    end_button = ttk.Button(root, text="End", command=root.destroy)
    end_button.pack()
    save_button = ttk.Button(root, text="Save to File", command=save_to_file)
    save_button.pack()
    root.after(1000, update_list)
    root.mainloop()