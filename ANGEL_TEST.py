def PermutationStep(num):
    # Convertir el número a lista de dígitos
    digits = list(str(num))
    n = len(digits)
    
    # Encontrar el primer dígito desde la derecha que es menor que su siguiente
    i = n - 2
    while i >= 0 and digits[i] >= digits[i + 1]:
        i -= 1
    
    # Si no encontramos tal dígito, no hay permutación mayor
    if i < 0:
        return -1
    
    # Encontrar el dígito más pequeño a la derecha que es mayor que digits[i]
    j = n - 1
    while digits[j] <= digits[i]:
        j -= 1
    
    # Intercambiar los dígitos
    digits[i], digits[j] = digits[j], digits[i]
    
    # Ordenar los dígitos a la derecha de i en orden ascendente
    left = digits[:i + 1]
    right = sorted(digits[i + 1:])
    
    # Combinar y convertir de nuevo a número
    result = int(''.join(left + right))
    
    return result

# Pruebas
test_cases = [41352, 11121, 999]
for test in test_cases:
    print(f"Número: {test} -> Siguiente permutación: {PermutationStep(test)}")