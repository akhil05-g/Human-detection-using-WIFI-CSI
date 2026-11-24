/*
 * =================================================================================
 * Project: Camera-Free Human Detection & Vital Sensing using Wi-Fi CSI
 * Target Hardware: ESP32 / ESP32-S3 / ESP32-WROOM-32 Dev Module
 * Framework: Arduino IDE (with ESP32 board package installed)
 * =================================================================================
 * How to Flash:
 * 1. Open Arduino IDE.
 * 2. Select Tools -> Board -> "ESP32 Dev Module" (or "ESP32S3 Dev Module").
 * 3. Update WIFI_SSID and WIFI_PASSWORD below with your Wi-Fi or Mobile Hotspot.
 * 4. Connect ESP32 to PC via USB cable, select the correct COM port under Tools -> Port.
 * 5. Click "Upload".
 * =================================================================================
 */

#include <WiFi.h>
#include <esp_wifi.h>
#include <esp_wifi_types.h>

// Enter your Wi-Fi Router or Phone Hotspot details
const char* WIFI_SSID = "YOUR_WIFI_NAME";
const char* WIFI_PASSWORD = "YOUR_WIFI_PASSWORD";

// Target IP to ping periodically to generate continuous Wi-Fi packet traffic
IPAddress gatewayIP;
unsigned long lastTrafficTime = 0;
const unsigned long TRAFFIC_INTERVAL_MS = 100; // 10 packets per second

// Callback triggered whenever a Wi-Fi packet with Channel State Information is received
void IRAM_ATTR wifi_csi_rx_cb(void *ctx, wifi_csi_info_t *info) {
    if (!info || !info->buf || info->len <= 0) return;

    // Packet header
    int8_t rssi = info->rx_ctrl.rssi;
    int len = info->len;
    int8_t *csi_buf = (int8_t *)info->buf;

    // Print CSI packet to Serial:
    // Format: CSI_DATA,<rssi>,<len>,<i0>,<q0>,<i1>,<q1>,...
    Serial.print("CSI_DATA,");
    Serial.print(rssi);
    Serial.print(",");
    Serial.print(len);

    // Stream IQ bytes
    for (int i = 0; i < len; i++) {
        Serial.print(",");
        Serial.print(csi_buf[i]);
    }
    Serial.println();
}

void setup() {
    Serial.begin(115200);
    delay(1000);
    Serial.println("\n[INIT] Starting ESP32 Wi-Fi CSI Collector...");

    // Connect to Wi-Fi station mode
    WiFi.mode(WIFI_STA);
    WiFi.begin(WIFI_SSID, WIFI_PASSWORD);

    Serial.print("[WIFI] Connecting to ");
    Serial.print(WIFI_SSID);
    int retry = 0;
    while (WiFi.status() != WL_CONNECTED && retry < 30) {
        delay(500);
        Serial.print(".");
        retry++;
    }

    if (WiFi.status() == WL_CONNECTED) {
        Serial.println("\n[WIFI] Connected successfully!");
        Serial.print("[WIFI] IP Address: ");
        Serial.println(WiFi.localIP());
        gatewayIP = WiFi.gatewayIP();
        Serial.print("[WIFI] Gateway IP: ");
        Serial.println(gatewayIP);
    } else {
        Serial.println("\n[WARN] Wi-Fi connection timed out. CSI collector will still listen in promiscuous mode.");
    }

    // Configure CSI capture parameters
    wifi_csi_config_t csi_config = {
        .lltf_en = true,
        .htltf_en = true,
        .stbc_htltf2_en = true,
        .ltf_merge_en = true,
        .channel_filter_en = true,
        .manu_scale = false,
        .shift = false,
    };

    esp_err_t err = esp_wifi_set_csi_config(&csi_config);
    if (err == ESP_OK) {
        esp_wifi_set_csi_rx_cb(wifi_csi_rx_cb, NULL);
        esp_wifi_set_csi(true);
        Serial.println("[CSI] Wi-Fi CSI Hardware Engine Activated!");
    } else {
        Serial.print("[ERROR] Failed to initialize CSI: ");
        Serial.println(err);
    }
}

void loop() {
    // Generate periodic Wi-Fi packet traffic so CSI frames are continuously captured
    unsigned long now = millis();
    if (now - lastTrafficTime >= TRAFFIC_INTERVAL_MS) {
        lastTrafficTime = now;
        if (WiFi.status() == WL_CONNECTED) {
            // Send a tiny UDP packet to gateway to stimulate Wi-Fi transmission
            WiFiUDP udp;
            udp.beginPacket(gatewayIP, 5005);
            udp.write((const uint8_t*)"PING", 4);
            udp.endPacket();
        }
    }
    delay(10);
}
