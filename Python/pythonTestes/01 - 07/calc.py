print("Calculadora em py")

CALC = input("Insira a a função desejada: Soma, Divisão, Multiplicação e Subtração: \n").lower()

num1 = int(input("Insira o primeiro valor: \n "))
num2 = int(input("Insira o segundo valor: \n "))

if CALC == "divisão" and num1 <= 0:
    print("Impossivel dividir por 0!")

if CALC == "soma":
    print(num1 + num2)
elif CALC == "divisão":
    print(num1 // num2)
elif CALC == "multiplicação":
    print(num1 * num2)
else:
    print(num1 - num2)