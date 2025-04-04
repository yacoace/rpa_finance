class TorreHanoi:
    def __init__(self):
        self.torres = {
            'A': [],
            'B': [],
            'C': []
        }
        self.movimientos = 0
    
    def inicializar_juego(self, n_discos):
        """Inicializa la torre A con n discos"""
        self.torres['A'] = list(range(n_discos, 0, -1))
        self.torres['B'] = []
        self.torres['C'] = []
        self.movimientos = 0
    
    def mostrar_torres(self):
        """Muestra el estado actual de las torres"""
        print("\nEstado actual de las torres:")
        altura_maxima = max(len(torre) for torre in self.torres.values())
        
        for nivel in range(altura_maxima - 1, -1, -1):
            for torre in ['A', 'B', 'C']:
                if nivel < len(self.torres[torre]):
                    print(f"[{self.torres[torre][nivel]}]", end='\t')
                else:
                    print("[ ]", end='\t')
            print()
        print(" A \t B \t C ")
        print(f"\nMovimientos realizados: {self.movimientos}")
    
    def mover_disco(self, origen, destino):
        """Intenta mover un disco de una torre a otra"""
        if not self.torres[origen]:
            print("\n¡Error! La torre de origen está vacía")
            return False
        
        if self.torres[destino] and self.torres[origen][-1] > self.torres[destino][-1]:
            print("\n¡Error! No puedes colocar un disco más grande sobre uno más pequeño")
            return False
        
        disco = self.torres[origen].pop()
        self.torres[destino].append(disco)
        self.movimientos += 1
        return True
    
    def verificar_victoria(self, n_discos):
        """Verifica si el juego se ha completado"""
        return len(self.torres['C']) == n_discos and \
               self.torres['C'] == list(range(n_discos, 0, -1))

def jugar_hanoi():
    """Función principal del juego"""
    print("¡Bienvenido a la Torre de Hanoi!")
    
    while True:
        try:
            n_discos = int(input("\nIngrese el número de discos (3-8): "))
            if 3 <= n_discos <= 8:
                break
            print("Por favor, ingrese un número entre 3 y 8")
        except ValueError:
            print("Por favor, ingrese un número válido")
    
    juego = TorreHanoi()
    juego.inicializar_juego(n_discos)
    min_movimientos = 2**n_discos - 1
    
    print(f"\nObjetivo: Mover todos los discos a la torre C")
    print(f"Número mínimo de movimientos posibles: {min_movimientos}")
    
    while not juego.verificar_victoria(n_discos):
        juego.mostrar_torres()
        
        while True:
            origen = input("\nTorre de origen (A/B/C) o 'Q' para salir: ").upper()
            if origen == 'Q':
                return
            if origen in ['A', 'B', 'C']:
                break
            print("Torre inválida. Use A, B, o C")
        
        while True:
            destino = input("Torre de destino (A/B/C): ").upper()
            if destino in ['A', 'B', 'C']:
                break
            print("Torre inválida. Use A, B, o C")
        
        juego.mover_disco(origen, destino)
    
    juego.mostrar_torres()
    print(f"\n¡Felicitaciones! Has completado el juego en {juego.movimientos} movimientos!")
    print(f"El mínimo posible era: {min_movimientos} movimientos")

if __name__ == "__main__":
    jugar_hanoi()