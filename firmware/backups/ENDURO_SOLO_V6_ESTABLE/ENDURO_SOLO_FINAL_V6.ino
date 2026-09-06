#include <EnduroSoloCore2V3.h>
#include <esp_system.h>

// ENDURO SOLO FINAL V6 - ESP32-C3 / Arduino ESP32 2.0.17 / iOS
// Pulsador normalmente abierto entre GPIO 4 y GND.
// Divisor de bateria: bateria+ -- 1 Mohm -- GPIO 3 -- 100 kohm -- GND.

const uint8_t PIN_BOTON = 4;
const uint8_t PIN_BATERIA = 3;

const uint32_t ANTIRREBOTE_MS = 35;
const uint32_t PULSACION_LARGA_MS = 2000;
const uint32_t INTERVALO_BATERIA_MS = 30000;

// 1 Mohm arriba y 100 kohm abajo: (1M + 100k) / 100k = 11.
// Se puede ajustar despues comparando el Monitor Serie con un multimetro.
const float FACTOR_DIVISOR = 11.0f;
// Calibrado con bateria cargada: lectura original 9.63 V, multimetro 4.08 V.
const float FACTOR_CALIBRACION = 0.424f;

BleKeyboard teclado("ENDURO SOLO V5", "Enduro Link", 100);

bool lecturaAnterior = HIGH;
bool estadoEstable = HIGH;
bool largaEnviada = false;
bool conexionAnterior = false;
uint32_t ultimoCambioMs = 0;
uint32_t inicioPulsacionMs = 0;
uint32_t ultimaBateriaMs = 0;

float leerVoltajeBateria() {
  uint32_t sumaMv = 0;
  const uint8_t muestras = 32;

  for (uint8_t i = 0; i < muestras; i++) {
    sumaMv += analogReadMilliVolts(PIN_BATERIA);
    delay(2);
  }

  const float mvPin = static_cast<float>(sumaMv) / muestras;
  return (mvPin / 1000.0f) * FACTOR_DIVISOR * FACTOR_CALIBRACION;
}

void enviarAtajoIOS(uint8_t tecla) {
  // Esta combinacion llega a Safari pero no escribe caracteres en otras apps.
  teclado.press(KEY_LEFT_CTRL);
  teclado.press(KEY_LEFT_ALT);
  teclado.press(KEY_LEFT_SHIFT);
  teclado.press(tecla);
  delay(20);
  teclado.releaseAll();
}

uint8_t porcentajeDesdeVoltaje(float voltaje) {
  // Aproximacion practica para una LiPo 1S bajo la carga liviana del ESP32.
  const float voltajes[] = {3.30f, 3.50f, 3.60f, 3.70f, 3.75f, 3.80f,
                            3.85f, 3.92f, 3.98f, 4.03f, 4.08f};
  const uint8_t porcentajes[] = {0, 5, 15, 30, 40, 50, 60, 70, 80, 90, 100};
  const uint8_t cantidad = sizeof(voltajes) / sizeof(voltajes[0]);

  if (voltaje <= voltajes[0]) return 0;
  if (voltaje >= voltajes[cantidad - 1]) return 100;

  for (uint8_t i = 1; i < cantidad; i++) {
    if (voltaje <= voltajes[i]) {
      const float proporcion =
          (voltaje - voltajes[i - 1]) / (voltajes[i] - voltajes[i - 1]);
      return static_cast<uint8_t>(
          porcentajes[i - 1] + proporcion * (porcentajes[i] - porcentajes[i - 1]) + 0.5f);
    }
  }

  return 0;
}

void informarBateria(bool enviarCodigoApp) {
  const float voltaje = leerVoltajeBateria();
  const uint8_t porcentaje = porcentajeDesdeVoltaje(voltaje);

  teclado.setBatteryLevel(porcentaje);

  if (teclado.isConnected() && enviarCodigoApp) {
    // "b" + 0..9 = 0..90%; "b" + "a" = 100%.
    uint8_t decena = (porcentaje + 5) / 10;
    if (decena > 10) decena = 10;
    teclado.write('b');
    teclado.write(decena < 10 ? static_cast<uint8_t>('0' + decena) : 'a');
  }

  Serial.print("Bateria: ");
  Serial.print(voltaje, 2);
  Serial.print(" V - ");
  Serial.print(porcentaje);
  Serial.println("%");
}

void enviarPaso() {
  if (!teclado.isConnected()) {
    Serial.println("Pulsacion sin conexion Bluetooth");
    return;
  }

  // La T simple es la señal confirmada como compatible con Safari en iOS.
  teclado.write('t');
  delay(25);
  informarBateria(true);
  Serial.println("PASO ENVIADO");
}

void mostrarOcultarTecladoIOS() {
  // Desactivado en V5: se usa un unico reporte HID para maxima estabilidad.
}

void setup() {
  Serial.begin(115200);
  delay(300);

  // Identidad Bluetooth fija y nueva para que iOS no reutilice los enlaces
  // defectuosos creados por las versiones anteriores.
  uint8_t macBaseNueva[6] = {0x02, 0x45, 0x4E, 0x44, 0x55, 0x35};
  esp_base_mac_addr_set(macBaseNueva);

  pinMode(PIN_BOTON, INPUT_PULLUP);
  analogReadResolution(12);
  analogSetPinAttenuation(PIN_BATERIA, ADC_0db);

  lecturaAnterior = digitalRead(PIN_BOTON);
  estadoEstable = lecturaAnterior;

  Serial.println();
  Serial.println("ENDURO SOLO FINAL V6");
  teclado.begin();
}

void loop() {
  const uint32_t ahora = millis();
  const bool lectura = digitalRead(PIN_BOTON);

  if (lectura != lecturaAnterior) {
    lecturaAnterior = lectura;
    ultimoCambioMs = ahora;
  }

  if ((ahora - ultimoCambioMs >= ANTIRREBOTE_MS) && lectura != estadoEstable) {
    estadoEstable = lectura;

    if (estadoEstable == LOW) {
      inicioPulsacionMs = ahora;
      largaEnviada = false;
      // Se envia al presionar, no al soltar, para conservar la precision.
      enviarPaso();
    } else {
      inicioPulsacionMs = 0;
      largaEnviada = false;
    }
  }

  if (estadoEstable == LOW && !largaEnviada &&
      ahora - inicioPulsacionMs >= PULSACION_LARGA_MS) {
    largaEnviada = true;
    mostrarOcultarTecladoIOS();
  }

  const bool conectado = teclado.isConnected();
  if (conectado && !conexionAnterior) {
    delay(250);
    informarBateria(false);
    ultimaBateriaMs = millis();
    Serial.println("Bluetooth conectado");
  }
  conexionAnterior = conectado;

  if (conectado && ahora - ultimaBateriaMs >= INTERVALO_BATERIA_MS) {
    informarBateria(false);
    ultimaBateriaMs = millis();
  }

  delay(5);
}
