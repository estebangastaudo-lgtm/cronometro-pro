# Enduro Solo V0.3

## Funciones

- Pulsación normal: envía `F12` y registra el paso.
- Pulsación mantenida durante 2 segundos: además envía `Eject` para mostrar u ocultar el teclado virtual de iOS.
- Batería: actualiza el servicio BLE y envía a la PWA el porcentaje redondeado cada 30 segundos.

## Instalación

1. Instalar `Enduro_BLE_Keyboard.zip` desde Arduino IDE: **Sketch > Include Library > Add .ZIP Library**.
2. Abrir `ENDURO_SOLO_V0_3.ino`.
3. Seleccionar **ESP32C3 Dev Module** y cargar el programa.
4. En el iPhone, olvidar el dispositivo Bluetooth `ENDURO SOLO` y vincularlo nuevamente. Es obligatorio porque cambió el descriptor HID.

## Conexiones

- Pulsador: `GPIO 4` a `GND`.
- Medición: batería positiva a `1 MΩ`, punto medio a `GPIO 3`, y `100 kΩ` desde el punto medio a `GND`.

La función de pulsación larga también envía inicialmente `F12` para no retrasar el cronometraje. Debe usarse con el cronómetro pausado; la PWA ignora el paso en ese estado.
