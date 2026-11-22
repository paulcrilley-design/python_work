# This is a sample Python script that demonstrates user input and output in Python as well as 
# 'Flag' usage to control the flow of the program

active = True
prompt = "\nType a message and I will repeat it: "
prompt +="\nType quit to exit. "
message = ""
while active:
    message = input(prompt)
    if message.lower() == 'quit':
        active = False
    else:
        print(message)
        print("I just repeated your message!")

