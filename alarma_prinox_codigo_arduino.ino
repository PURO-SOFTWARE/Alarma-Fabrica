const int relePin = 2;
int valor_1 = 0;

void setup() {
  pinMode(relePin, OUTPUT);
  Serial.begin(9600);
}

void loop() {
  if (Serial.available() > 0) {
    valor_1 = Serial.parseInt();
    if (valor_1 == 0) {
      digitalWrite(relePin, HIGH); // Activar el relé
      delay(5000); // Mantener el relé activo durante 5 segundos
      digitalWrite(relePin, LOW); // Desactivar el relé
    }
  }
}
