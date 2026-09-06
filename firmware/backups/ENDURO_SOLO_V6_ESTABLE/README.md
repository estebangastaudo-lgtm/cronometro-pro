# ENDURO SOLO V6 ESTABLE — BACKUP

Respaldo de la versión comprobada en la ESP32-C3 antes de integrar ESP-NOW.

## Archivos

- `ENDURO_SOLO_FINAL_V6.ino`: firmware estable.
- `Enduro_Solo_Core2_V3.zip`: biblioteca Bluetooth exacta requerida por el firmware.

## Configuración de Arduino IDE

- Paquete de placas: **esp32 by Espressif Systems 2.0.17**.
- Placa: **ESP32C3 Dev Module**.
- USB CDC On Boot: **Enabled**.
- Upload Speed: **115200**.
- Erase All Flash Before Sketch Upload: **Disabled**.
- Monitor Serie: **115200 baud**.

## Cableado comprobado

- Pulsador normalmente abierto: **GPIO 4 ↔ GND**.
- Divisor de batería: **batería+ → 1 MΩ → GPIO 3 → 100 kΩ → GND**.
- Calibración comprobada: batería real **4,08 V**, mostrada como **4,08 V / 100 %**.

## Restauración

1. Instalar `Enduro_Solo_Core2_V3.zip` desde **Programa → Incluir librería → Añadir biblioteca .ZIP**.
2. Abrir `ENDURO_SOLO_FINAL_V6.ino`.
3. Seleccionar las opciones anteriores y cargar el sketch.

Esta versión es el punto de retorno estable. No modificar estos archivos.
