# test_integracion.py

import unittest
from Interfaz_usuario2 import crear_interfaz
from Iniciar_Sesion import iniciar_sesion

class TestIntegracion(unittest.TestCase):
    def test_flujo_completo(self):
        # Verificar que el flujo completo desde la interfaz de usuario hasta el inicio de sesión funcione correctamente
        crear_interfaz()
        iniciar_sesion('usuario', 'contraseña')

if __name__ == '__main__':
    unittest.main()
