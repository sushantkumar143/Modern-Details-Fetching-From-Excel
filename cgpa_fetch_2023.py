import csv
import private
import win32com.client as wincom

def speak():
    speak = wincom.Dispatch("SAPI.SpVoice")
    text = "Welcome to ROBOT PROJECT that is developed by Sushant kumar"
    speak.speak(text)
speak()

def speak_1(a):
        speak = wincom.Dispatch("SAPI.SpVoice")
        speak.speak(a)

file = open(r"C:\Users\susha\OneDrive\Desktop\Python_project\cgpa.csv", "r", newline="")
obj = csv.reader(file)


def display(data, n):
    print("Registration Number            :     ",data[0])
    print("Name                           :     ",data[1])
    print("Father's Name                  :     ",data[2])
    print("Mother's Name                  :     ",data[3])
    print("Course Details                 :     ",data[4])
    print("Program Type                   :     ",data[5])
    print("Batch Year                     :     ",data[6])
    print("Gender                         :     ",data[7])
    print("Program Duration               :     ",data[8])
    print("Country                        :     ",data[9])
    print("State                          :     ",data[10])
    print("Section                        :     ",data[11])
    print("-------------------------------------------")
    print("| CGPA                         :     ",data[12],"|")
    print("| Rank                         :     ",n,"|")
    print("-------------------------------------------")
    print("Total Coarses                  :     ",data[13])
    print("Passed Coarses                 :     ",data[14])

def registration():
    reg = input("Enter Registration Number: ")
    return reg

for i in range(1,4):
    pwd = input("Enter the Password: ")
    if pwd == private.pwd:
        speak_1("Congratulation, you now have the access of students details")
        while True:
            reg = registration()
            chk = 0
            rank = 0
            for data in obj:
                rank += 1
                if reg in data:
                    # print(data)
                    display(data, rank)
                    speak_1("Name: "+str(data[1]))
                    speak_1("CGPA: " + str(data[12]))
                    speak_1("Rank: " + str(rank))
                    chk = 1
                    break

            if chk==0:
                print("\nRegistration Number ",reg, " Not Found.\n")

            ask = input("\nDo you want continue (y/n): ")
            if ask=="n":
                speak_1("Thank you dear")
                break
        break

    elif i==3:
        print("               -------------------------------")
        print("               |     !! ACCESS DENIED !!      |")
        print("               |     Try After Some Time      |")
        print("               -------------------------------")
        speak_1("Access Denied, Please after some time.")
        break

    else:
        print("!! Wrong Password")
        print("Warning: Only ", 3-i , " times left otherwise access will denied\n")
        speak_1("Only " + str(3-i) + "attempt left")


file.close()