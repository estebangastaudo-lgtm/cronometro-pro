#include <EnduroSoloCore2V3.h>
#include <WiFi.h>
#include <esp_now.h>
#include <esp_system.h>
#include <esp_wifi.h>

// ENDURO SOLO + LINK FINAL V7
// ESP32-C3 / Arduino ESP32 2.0.17 / iOS / ESP-NOW
//
// Una unica pulsacion:
//   1. Envia el evento por ESP-NOW a la Base DevKit de la camioneta.
//   2. Envia la tecla "t" por Bluetooth al telefono de la moto.
//
// Pulsador normalmente abierto entre GPIO 4 y GND.
// Divisor de bateria: bateria+ -- 1 Mohm -- GPIO 3 -- 100 kohm -- GND.

const uint8_t PIN_BOTON = 4;
const uint8_t PIN_BATERIA = 3;

const uint32_t ANTIRREBOTE_MS = 35;
const uint32_t PULSACION_LARGA_MS = 2000;
const uint32_t INTERVALO_BATERIA_MS = 30000;

// Divisor y calibracion comprobados en Enduro Solo V6.
const float FACTOR_DIVISOR = 11.0f;
const float FACTOR_CALIBRACION = 0.424f;

// MAC real de la Base DevKit de Enduro Link.
uint8_t macBase[] = {0x20, 0xE7, 0xC8, 0x5A, 0xC8, 0x0C};

// Formato de paquete ya probado con la Base DevKit.
struct ELP_Frame {
  uint8_t version;
  uint8_t tipo;
  uint8_t idNodo;
  uint8_t flags;
  uint32_t numeroEvento;
  uint32_t tiempoNodoMs;
  uint8_t bateria;
  uint16_t crc;
};

// Se conserva el mismo nombre y la misma identidad Bluetooth de la V6 para
// que el iPhone use el enlace estable que ya tiene guardado.
BleKeyboard teclado("ENDURO SOLO V5", "Enduro Link", 100);

bool lecturaAnterior = HIGH;
bool estadoEstable = HIGH;
bool largaEnviada = false;
bool conexionAnterior = false;
bool espNowListo = false;

uint32_t ultimoCambioMs = 0;
uint32_t inicioPulsacionMs = 0;
uint32_t ultimaBateriaMs = 0;
uint32_t numeroEvento = 0;

float voltajeBateriaActual = 0.0f;
uint8_t porcentajeBateriaActual = 100;

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

uint8_t porcentajeDesdeVoltaje(float voltaje) {
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
          porcentajes[i - 1] +
          proporcion * (porcentajes[i] - porcentajes[i - 1]) + 0.5f);
    }
  }

  return 0;
}

void informarBateria(bool enviarCodigoApp) {
  voltajeBateriaActual = leerVoltajeBateria();
  porcentajeBateriaActual = porcentajeDesdeVoltaje(voltajeBateriaActual);

  teclado.setBatteryLevel(porcentajeBateriaActual);

  if (teclado.isConnected() && enviarCodigoApp) {
    // "b" + 0..9 = 0..90%; "b" + "a" = 100%.
    uint8_t decena = (porcentajeBateriaActual + 5) / 10;
    if (decena > 10) decena = 10;
    teclado.write('b');
    teclado.write(decena < 10 ? static_cast<uint8_t>('0' + decena) : 'a');
  }

  Serial.print("Bateria: ");
  Serial.print(voltajeBateriaActual, 2);
  Serial.print(" V - ");
  Serial.print(porcentajeBateriaActual);
  Serial.println("%");
}

// Firma correspondiente al paquete ESP32 by Espressif Systems 2.0.17.
void alEnviar(const uint8_t *mac, esp_now_send_status_t estado) {
  (void)mac;

  if (estado == ESP_NOW_SEND_SUCCESS) {
    Serial.println("LINK: envio confirmado por la Base");
  } else {
    // No se bloquea ni se reinicia: al volver al alcance, la siguiente
    // pulsacion se envia normalmente sin reconexion manual.
    Serial.println("LINK: Base fuera de alcance");
  }
}

void iniciarEnduroLink() {
  WiFi.mode(WIFI_STA);
  delay(250);

  if (esp_now_init() != ESP_OK) {
    Serial.println("LINK: ERROR iniciando ESP-NOW");
    return;
  }

  esp_now_register_send_cb(alEnviar);

  esp_now_peer_info_t peerInfo = {};
  memcpy(peerInfo.peer_addr, macBase, 6);
  peerInfo.channel = 0;
  peerInfo.encrypt = false;
  peerInfo.ifidx = WIFI_IF_STA;

  if (esp_now_add_peer(&peerInfo) != ESP_OK) {
    Serial.println("LINK: ERROR agregando la Base");
    return;
  }

  espNowListo = true;
  Serial.println("LINK: listo para enviar a la camioneta");
}

void enviarPasoEnduroLink(uint32_t instantePulsacionMs) {
  if (!espNowListo) {
    Serial.println("LINK: no disponible");
    return;
  }

  numeroEvento++;

  ELP_Frame paquete = {};
  paquete.version = 1;
  paquete.tipo = 1;
  paquete.idNodo = 1;
  paquete.flags = 0;
  paquete.numeroEvento = numeroEvento;
  paquete.tiempoNodoMs = instantePulsacionMs;
  paquete.bateria = porcentajeBateriaActual;
  paquete.crc = 0;

  const esp_err_t resultado = esp_now_send(
      macBase,
      reinterpret_cast<const uint8_t *>(&paquete),
      sizeof(paquete));

  if (resultado == ESP_OK) {
    Serial.print("LINK: evento ");
    Serial.print(numeroEvento);
    Serial.println(" enviado");
  } else {
    Serial.print("LINK: ERROR al iniciar envio: ");
    Serial.println(resultado);
  }
}

void enviarPasoSolo() {
  if (!teclado.isConnected()) {
    Serial.println("SOLO: pulsacion sin conexion Bluetooth");
    return;
  }

  // La T simple es la senal confirmada como compatible con Safari en iOS.
  teclado.write('t');
  delay(25);
  informarBateria(true);
  Serial.println("SOLO: paso enviado");
}

void enviarPasoCompleto(uint32_t instantePulsacionMs) {
  // Los dos caminos son independientes. Un fallo en uno no cancela el otro.
  enviarPasoEnduroLink(instantePulsacionMs);
  enviarPasoSolo();
}

void mostrarOcultarTecladoIOS() {
  // Deliberadamente desactivado: un unico reporte HID mantiene estable iOS.
}

void setup() {
  Serial.begin(115200);
  delay(300);

  // Misma identidad base de la V6 para conservar el emparejamiento Bluetooth.
  uint8_t macBaseNueva[6] = {0x02, 0x45, 0x4E, 0x44, 0x55, 0x35};
  esp_base_mac_addr_set(macBaseNueva);

  pinMode(PIN_BOTON, INPUT_PULLUP);
  pinMode(PIN_BATERIA, INPUT);
  analogReadResolution(12);
  analogSetPinAttenuation(PIN_BATERIA, ADC_0db);

  lecturaAnterior = digitalRead(PIN_BOTON);
  estadoEstable = lecturaAnterior;

  Serial.println();
  Serial.println("ENDURO SOLO + LINK FINAL V7");

  iniciarEnduroLink();
  teclado.begin();
  informarBateria(false);
  ultimaBateriaMs = millis();
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
      // El instante se captura al presionar, antes de cualquier envio o lectura.
      enviarPasoCompleto(ahora);
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
    Serial.println("SOLO: Bluetooth conectado");
  }
  conexionAnterior = conectado;

  // La bateria se actualiza aunque Bluetooth este desconectado, para que el
  // siguiente paquete de Enduro Link lleve un valor reciente.
  if (ahora - ultimaBateriaMs >= INTERVALO_BATERIA_MS) {
    informarBateria(false);
    ultimaBateriaMs = millis();
  }

  delay(5);
}
