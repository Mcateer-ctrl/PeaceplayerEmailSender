import csv
import time
import sys
import os
import time
import smtplib
import threading
import logging
from email.message import EmailMessage
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk

email_address = "josh.mcateer50@gmail.com"
email_password = "mxxcbnxzxkhbrvzc"

is_canceled = False

logging.basicConfig(filename='email_sender.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def browse_file():
    while True:
        file_path = filedialog.askopenfilename()
        if not file_path:
            break  # If the user cancels the file dialog, exit the loop

        # Extract the filename from the file path
        filename = os.path.basename(file_path)

        # Check if the selected file is a CSV file
        if not file_path.lower().endswith(".csv"):
            messagebox.showerror("Invalid file", "Please select a CSV file.")
            continue  # Re-open the file dialog if the file is not a CSV file

        csv_file_entry.delete(0, tk.END)
        csv_file_entry.insert(0, file_path)

        if messagebox.askyesno("Confirmation", f"You have selected '{filename}'. Is this correct?"):
            # Perform any necessary actions based on the filename here
            break  # Exit the loop if the user selects "Yes"
        else:
            continue  # Re-open the file dialog if the user selects "No"

def cancel_emails():
    global is_canceled
    is_canceled = True

def send_emails():
    # Get the CSV file path
    file_path = csv_file_entry.get()

    global is_canceled

    logging.info("Email sending started.")


    # Check if the CSV file name contains specific strings
    if "north" in file_path:
        # Change data variables for "north"
        logging.info("north spreadsheet chosen.")
        pass  # Replace this line with your logic
    elif "south" in file_path:
        # Change data variables for "south"
        logging.info("south spreadsheet chosen.")
        pass  # Replace this line with your logic
    elif "east" in file_path:
        # Change data variables for "east"
        logging.info("east spreadsheet chosen.")
        pass  # Replace this line with your logic
    elif "west" in file_path:
        # Change data variables for "west"
        logging.info("west spreadsheet chosen.")
        pass  # Replace this line with your logic

    # Open the CSV file
    try:
        with open(file_path, newline='') as csvfile:
            # Create a CSV reader object
            csvreader = csv.reader(csvfile, delimiter=',')

            # Skip the first row
            next(csvreader)

            # Loop through each row in the CSV file
            email_count = 0
            for row in csvreader:
                if is_canceled:
                    break
                # Extract the required data
                parentName = row[0]
                childName = row[2]
                Email = row[1]
                commentOption = row[4]

                email_count += 1

                # Personalized email message
                if commentOption == "Natural Leader/ Confident":
                    commentParagraph = f"Through our programmes we have noticed how much of a natural leader {childName} is and how their confidence and skills have grown over the years."
                elif commentOption == "Jokester/ Funny/ Friendly":
                    commentParagraph = f"Through our programmes we have seen {childName}'s friendly and funny personality grow and shine, and their ability to include and look out for others."
                elif commentOption == "Kind, Gentle and Nice":
                    commentParagraph = f"{childName} is a kind and caring member of the team, always polite and thoughtful to those around them."
                elif commentOption == "Insightful/ Always brings up interesting ideas and topics":
                    commentParagraph = f"{childName} is always engaged during our sessions, keen and eager to learn from others and always brings up interesting ideas and topics."
                elif commentOption == "Independent/ likes to experiment and explore during games and team time":
                    commentParagraph = f"{childName} comes to every session with high energy, ready to engage, and is confident and independent, always likes to experiment and explore during games and team time."
                else:
                    commentParagraph = ""

                message = f"Dear {parentName},\n\nWe at PeacePlayers believe that sport has the power to bring people together, break down barriers and create positive change in our communities. Throughout the past number of years {childName} has been a valuable member of PeacePlayers and has been involved in our after school programme BDP (formerly CCL) providing cross-community basketball training, matches and community relations discussions.\n\n{commentParagraph}\n\nWe're launching a monthly giving campaign this week to help build important funds that will ensure the continued running of and wider impact of our programmes. By donating just a small amount each month, you can make a big difference in the lives of young people in Northern Ireland who are learning valuable skills like leadership, teamwork, and conflict resolution through basketball, just like {childName}. So, we're asking you to join us in supporting PeacePlayers Northern Ireland by signing up and making a monthly donation. For the small price of a cup of coffee a month, your donation will help PeacePlayers continue to deliver programming to both Catholic and Protestant youth in Northern Ireland which we believe is essential for building lasting relationships and understanding between communities that have historically been divided. Your generosity will help ensure that the organisation can continue its important work and make a meaningful impact on the lives of young people in Northern Ireland and towards a shared future.\n\nThank you for your support.\n\nSincerely,\nPeacePlayers Northern Ireland"
                try:
                    # create email
                    msg = EmailMessage()
                    msg['Subject'] = "PeacePlayers Northern Ireland - Monthly Giving Campaign"
                    msg['From'] = email_address
                    msg['To'] = Email
                    msg.set_content(message)

                    # send email
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email_address, email_password)
                        smtp.send_message(msg)

                    # Add a status message to the display box
                    display_box.insert(tk.END, f"Email sent successfully to {Email}\n")
                    logging.info(f"Email sent successfully to {Email}")
                    display_box.see(tk.END)
                    root.update_idletasks()  # Force the GUI to update the display box
                except Exception as e:
                    # Add an error message to the display box
                    display_box.insert(tk.END, f"Error occurred while sending email to {Email}: {e}\n")
                    logging.error(f"Error occurred while sending email to {Email}: {e}")
                    display_box.see(tk.END)
                    root.update_idletasks()  # Force the GUI to update the display box
                # Wait for 1 minute after sending every 20 emails
                if email_count % 20 == 0:
                    time.sleep(60)
        # Show a success message
        if is_canceled:
            display_box.insert(tk.END, "Email sending process canceled.\n")
            logging.info("Email sending process canceled.")
            display_box.see(tk.END)
            root.update_idletasks()  # Force the GUI to update the display box
            is_canceled = False
        else:
            messagebox.showinfo("Success", "Emails sent successfully!")
            logging.info("Email sending completed.")
    except IOError:
        messagebox.showerror("Error", "Could not read CSV file")
        logging.error("Could not read CSV file")

def start_sending_emails():
    email_thread = threading.Thread(target=send_emails)
    email_thread.start()


def resource_path(relative_path):
    """Get the correct resource path for the running mode."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def show_splash_screen():
    root.withdraw()  # Hide the main window
    splash = tk.Toplevel()
    splash.title("Peaceplayers")
    splash.geometry("600x400")
    splash.config(bg="white")

    
    # Load the logo image
    image_path = resource_path("PeacePlayers-FB-Link-Preview.png")
    image = Image.open(image_path)
    logo = ImageTk.PhotoImage(image)
    logo_label = ttk.Label(splash, image=logo, background="white")
    logo_label.image = logo  # Store the image as an attribute of the label
    logo_label.pack(pady=20)

    # Create a canvas for the loading bar
    canvas = tk.Canvas(splash, width=300, height=20, bg="white", highlightthickness=0)
    canvas.pack(pady=10)

    def close_splash():
        splash.destroy()
        root.deiconify()  # Show the main window after destroying the splash screen

    def animate_loading_bar():
        for i in range(1, 101):  # Loading percentage from 1 to 100
            canvas.delete("all")  # Clear the canvas
            canvas.create_rectangle(0, 0, i * 3, 20, fill="blue", outline="")
            canvas.create_text(150, 10, text=f"{i}%", font=("Arial", 12, "bold"), fill="white")
            splash.update()  # Update the splash screen to show the changes
            time.sleep(0.01)  # Pause for 10 milliseconds

        close_splash()  # Close the splash screen when loading is complete

    splash.after(1, animate_loading_bar)  # Start animating the loading bar

def clear_display_box():
    display_box.delete(1.0, tk.END)




root = tk.Tk()
root.title("PeacePlayers Email Sender")
style = ttk.Style()
style.configure("TButton", font=("Arial", 12))
style.configure("TLabel", font=("Arial", 12))
style.configure("TEntry", font=("Arial", 12))

csv_file_label = ttk.Label(root, text="CSV File:")
csv_file_entry = ttk.Entry(root, width=50)
csv_file_browse_button = ttk.Button(root, text="Browse", command=browse_file)
send_emails_button = ttk.Button(root, text="Send Emails", command=start_sending_emails)
display_box_label = ttk.Label(root, text="Email Status:")
display_box = scrolledtext.ScrolledText(root, width=60, height=10)
cancel_emails_button = ttk.Button(root, text="Cancel", command=cancel_emails)
clear_display_button = ttk.Button(root, text="Clear Display", command=clear_display_box)



csv_file_label.grid(row=0, column=0, padx=10, pady=10)
csv_file_entry.grid(row=0, column=1, padx=10, pady=10)
csv_file_browse_button.grid(row=0, column=2, padx=10, pady=10)
send_emails_button.grid(row=1, column=1, padx=10, pady=10)
display_box_label.grid(row=2, column=0, padx=10, pady=10)
display_box.grid(row=2, column=1, padx=10, pady=10)
cancel_emails_button.grid(row=1, column=2, padx=10, pady=10)
clear_display_button.grid(row=3, column=1, padx=10, pady=10)


# Show the animated splash screen
show_splash_screen()
root.mainloop()
