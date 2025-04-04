let torres = {
    'A': [],
    'B': [],
    'C': []
};

let movimientos = 0;
let torreSeleccionada = null;
let numDiscos = 3;

function iniciarJuego() {
    // Reiniciar variables
    torres = {
        'A': [],
        'B': [],
        'C': []
    };
    movimientos = 0;
    torreSeleccionada = null;
    
    // Obtener número de discos seleccionado
    numDiscos = parseInt(document.getElementById('numDiscos').value);
    
    // Crear discos iniciales
    for (let i = numDiscos; i > 0; i--) {
        torres['A'].push(i);
    }
    
    actualizarInterfaz();
    document.getElementById('mensaje').textContent = "¡Juego iniciado! Selecciona una torre para mover un disco";
    document.getElementById('movimientos').textContent = "Movimientos: 0";
}

function actualizarInterfaz() {
    // Limpiar torres
    ['A', 'B', 'C'].forEach(torre => {
        const torreElement = document.getElementById(`torre${torre}`);
        // Mantener solo la base y el poste
        while (torreElement.children.length > 2) {
            torreElement.removeChild(torreElement.lastChild);
        }
    });
    
    // Agregar discos a cada torre
    ['A', 'B', 'C'].forEach(torre => {
        const torreElement = document.getElementById(`torre${torre}`);
        torres[torre].forEach(disco => {
            const discoElement = document.createElement('div');
            discoElement.className = 'disco';
            discoElement.style.width = `${disco * 30}px`;
            discoElement.textContent = disco;
            torreElement.appendChild(discoElement);
        });
    });
}

function seleccionarTorre(torre) {
    if (torreSeleccionada === null) {
        // Primera selección
        if (torres[torre].length === 0) {
            document.getElementById('mensaje').textContent = "¡Torre vacía! Selecciona otra torre";
            return;
        }
        torreSeleccionada = torre;
        document.getElementById(`torre${torre}`).classList.add('seleccionada');
        document.getElementById('mensaje').textContent = "Ahora selecciona la torre destino";
    } else {
        // Segunda selección (torre destino)
        if (torreSeleccionada === torre) {
            // Deseleccionar si se hace clic en la misma torre
            document.getElementById(`torre${torre}`).classList.remove('seleccionada');
            torreSeleccionada = null;
            document.getElementById('mensaje').textContent = "Selecciona una torre para mover un disco";
            return;
        }
        
        // Intentar mover el disco
        if (moverDisco(torreSeleccionada, torre)) {
            movimientos++;
            document.getElementById('movimientos').textContent = `Movimientos: ${movimientos}`;
            
            // Verificar victoria
            if (verificarVictoria()) {
                document.getElementById('mensaje').textContent = 
                    `¡Felicitaciones! Has completado el juego en ${movimientos} movimientos. ` +
                    `Mínimo posible: ${Math.pow(2, numDiscos) - 1}`;
            }
        }
        
        // Limpiar selección
        document.getElementById(`torre${torreSeleccionada}`).classList.remove('seleccionada');
        torreSeleccionada = null;
        actualizarInterfaz();
    }
}

function moverDisco(origen, destino) {
    if (torres[origen].length === 0) {
        document.getElementById('mensaje').textContent = "¡Error! La torre de origen está vacía";
        return false;
    }
    
    if (torres[destino].length > 0 && 
        torres[origen][torres[origen].length - 1] > torres[destino][torres[destino].length - 1]) {
        document.getElementById('mensaje').textContent = 
            "¡Error! No puedes colocar un disco más grande sobre uno más pequeño";
        return false;
    }
    
    torres[destino].push(torres[origen].pop());
    document.getElementById('mensaje').textContent = "Movimiento realizado. Continúa jugando";
    return true;
}

function verificarVictoria() {
    return torres['C'].length === numDiscos && 
           torres['C'].every((disco, index) => disco === numDiscos - index);
}

// Inicializar el juego cuando se carga la página
document.addEventListener('DOMContentLoaded', () => {
    iniciarJuego();
}); 