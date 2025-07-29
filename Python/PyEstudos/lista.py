
frutas = ["banana", "uva", "pera", "manga", "macaxeira", "morango", "jabuticaba", "mamao"]
numeros = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
ambos = [5, "bananas", 10, "peras", 15, "maças", 2, "kiwi"] 


print(frutas[0], "\n<- Essa é a primeira fruta")
print("\n_____")
print(frutas[1], "<- Essa é a segunda fruta")
print("\n_____")
print(frutas[2], "<- Essa é a terceira fruta")
print("\n_____")
print(frutas[3], "<- Essa é a quarta fruta")
print("\n_____")
print(frutas[-1], "<- Essa é a ultima fruta")
print("\n_____")
print(frutas[-2], "<- Essa é a penultima fruta")
print("\n_____")

toda_lista = print("Essa é a lista de todas as frutas") 
print(frutas)

cont = input("Aperte Enter para seguir para a lista de numeros")


print(numeros[0], "<- Esse é o primeiro numero")
print(numeros[1], "<- Esse é o segundo numero\n")

print(numeros[0] + numeros[1], "<- Essa é a soma do primeiro numero mais o segundo numero")
print(numeros[9] - numeros[4], "<- Essa é a subtração do ultimo numero com o quinto numero")

continuação = str(input("Caso você queira entender como funciona a soma das posições de uma lista, responda com X, caso queria continuar, aperte Enter")).lower()

if continuação == 'x':
    print("As contas em listas funcionam de uma maneira simples, é nescessario fazer a soma com as posições das listas\n Por exemplo:\n numeros[0] + numeros[1]\n isso vai fazer com que sejam somadas o primeiro valor com o segundo valor da lista!\n")
    input("Aperte Enter para continuar")


print(ambos)

print("\n A lista acima utiliza inteiros e strings")