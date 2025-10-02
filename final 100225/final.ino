// --- Pin Definitions ---
const int relayPins[] = {10, 11, 22, 23, 16};
const int relayCount = sizeof(relayPins) / sizeof(relayPins[0]);

const int inputPins[] = {1, 2, 3, 17}; // Job Sensors 1, 2, 3, and Door Sensor
const int inputCount = sizeof(inputPins) / sizeof(inputPins[0]);

// --- Timer for Periodic Updates ---
unsigned long previousMillis = 0;
const long interval = 1000; // Interval to send status (1000ms = 1 second)

// --- Handshake State Variable ---
bool gui_is_ready = false;

//------------------------------------------------------------------------------------

void setup() {
  Serial.begin(115200);

  // A short delay to ensure serial port is stable on boot
  delay(1000); 
  // Serial.println("ESP32 Booted. Waiting for GUI..."); // <-- THIS LINE IS REMOVED

  // Configure Relay Pins for Active-HIGH operation (LOW = OFF)
  for (int i = 0; i < relayCount; i++) {
    pinMode(relayPins[i], OUTPUT);
    digitalWrite(relayPins[i], LOW);
  }

  // Configure Input Pins with internal pull-up resistor.
  for (int i = 0; i < inputCount; i++) {
    pinMode(inputPins[i], INPUT_PULLUP);
  }
}

//------------------------------------------------------------------------------------

void loop() {
  // --- Handshake Logic ---
  // The board will do nothing until it receives a "PING" from the GUI.
  if (!gui_is_ready) {
    if (Serial.available() > 0) {
      String command = Serial.readStringUntil('\n');
      command.trim();
      if (command == "PING") {
        Serial.println("READY"); // Send confirmation back to the GUI
        gui_is_ready = true;
      }
    }
  } else {
    // --- NORMAL OPERATION (This runs only AFTER handshake is complete) ---

    // Part 1: Handle commands sent FROM the Python GUI
    if (Serial.available() > 0) {
      String command = Serial.readStringUntil('\n');
      command.trim();
      processCommand(command);
    }

    // Part 2: Send periodic status TO the Python GUI
    unsigned long currentMillis = millis();
    if (currentMillis - previousMillis >= interval) {
      previousMillis = currentMillis;
      
      for (int i = 0; i < inputCount; i++) {
        int currentState = digitalRead(inputPins[i]); // Reads 0 for LOW (Active), 1 for HIGH (Inactive)
        int reportedState = currentState;
        
        // Only flip the state for the first 3 sensors (the job sensors)
        if (i < 3) {
          reportedState = 1 - currentState; // Inverts 0->1 and 1->0
        }
        
        Serial.print("INPUT:");
        Serial.print(i);
        Serial.print(":");
        Serial.println(reportedState);
      }
    }
  }
}

//------------------------------------------------------------------------------------

void processCommand(String cmd) {
  cmd.toUpperCase();

  if (cmd.startsWith("RSET")) {
    int relayIndex, state;
    sscanf(cmd.c_str(), "RSET %d %d", &relayIndex, &state);
    if (relayIndex >= 0 && relayIndex < relayCount) {
      digitalWrite(relayPins[relayIndex], state == 1 ? HIGH : LOW);
    }
  } else if (cmd.startsWith("PULSE")) {
    int relayIndex, duration;
    sscanf(cmd.c_str(), "PULSE %d %d", &relayIndex, &duration);
    if (relayIndex >= 0 && relayIndex < relayCount) {
      digitalWrite(relayPins[relayIndex], HIGH);
      delay(duration);
      digitalWrite(relayPins[relayIndex], LOW);
    }
  }
}
