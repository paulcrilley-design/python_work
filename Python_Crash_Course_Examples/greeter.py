# More examples of user input and output in Python.

prompt = "If you enter your name the responses will be personalized: "
prompt += " Please enter your name: "
name = input(prompt)
print(f"Hello, {name}!")
print("It's nice to meet you!")

age = input("Please enter your age: ")
age = int(age)
print(f"You are {age} years old.")
print("Thank you for providing your information!")
if age >= 18:
    print("You are an adult.")
else:
    print("You are a legally a minor, so there will be some online usage restrictions.")

    



