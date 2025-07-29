# Repetição é uma das paradas mais chatas que podem ser feitas no python, o autor deste github e deste script meia boca aqui é o hater numero 1 de repetições :D

# Caso você goste de repetições acho que deveria ser tratar beijos de luz

a = int(input("Coloque aqui o valor de A\n"))
b = int(input("Coloque aqui o valor de B\n"))
c = int(input("Coloque aqui o valor de C\n"))

def soma():
    return a + b + c

for O in range(21):
    resultado = soma() + O
    print(f"\n {a} + {b} + {c} = {resultado}\n")

d = int(input("Coloque aqui o valor de D \n"))

for D in range(11):
    d += 1
    print(d)

i = int(input("Coloque o valor de i \n"))

if i >= 20:
    print(f"Soma = {soma()}")
    print(f"{i} + Soma = {i + soma()}")

elif i <= 10:
    print(f"Soma = {soma()}")
    print(f"{i} + {a} + {b} + {i} = {i + a + b + i}")

else:
    print("uai")

print( i + soma(), "Está sendo somada a função + o valor de i" )


t = int(input("Coloque aqui o valor de t\n"))
y = int(input("Coloque aqui o valor de y\n"))
h = int(input("Coloque aqui o valor de h\n"))

print("Os valores são ", t, y, h)

def YYY():
    return t + y + h

print(YYY() - 1, "porque os valores anteriores foram somados entre si e subtraidos por 1")