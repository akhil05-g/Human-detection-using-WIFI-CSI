# 📡 Wi-Fi CSI Camera-Free Human Detection & Vital Signs Project

## 1. What Hardware to Ask Your Faculty For
Ask your faculty or department lab technician for:

1. **Microcontroller Board (Any of these will work):**
   * **ESP32 Dev Module** (ESP32-WROOM-32) — *Most common in college labs (~$4)*
   * **ESP32-S3 DevKit** — *Newer dual-core version (~$8)*
   * **ESP32-C6** — *Wi-Fi 6 version*
2. **USB Data Cable:**
   * Micro-USB or USB-C (depending on your ESP32 board's port).
   * **Note:** Ensure it is a **data cable** (capable of data transfer), not a cheap power-only charging cable.
3. **Wi-Fi Network:**
   * Any standard Wi-Fi router OR simply your **Mobile Hotspot** enabled on your smartphone.

---

## 2. Hardware Connection Diagram

```
+-------------------+                   +------------------+                   +------------------+
|   Wi-Fi Router    | ~~~~~~~~~~~~~~~>  |   Human Body     | ~~~~~~~~~~~~~~~>  |   ESP32 Board    |
|   (or Phone AP)   |   Radio Waves     | (Reflects Waves) |  Perturbed Waves  |  (CSI Receiver)  |
+-------------------+                   +------------------+                   +------------------+
                                                                                         |
                                                                                USB Data Cable
                                                                                         v
                                                                               +------------------+
                                                                               |   Your Laptop    |
                                                                               | (Python + WebUI) |
                                                                               +------------------+
```

---

## 3. Step-by-Step: How to Flash the ESP32 (Takes 2 Minutes)

1. **Install Arduino IDE** (if not already installed).
2. In Arduino IDE, go to **Tools -> Board -> Boards Manager**, search for `esp32` by *Espressif Systems*, and click **Install**.
3. Open the firmware file:
   `c:\Users\Akhil\OneDrive\Desktop\ruview\Human_Detection_CSI\firmware_arduino\ESP32_CSI_Collector.ino`
4. In lines 19-20, set your Wi-Fi name and password:
   ```cpp
   const char* WIFI_SSID = "YOUR_HOTSPOT_NAME";
   const char* WIFI_PASSWORD = "YOUR_PASSWORD";
   ```
5. Plug your ESP32 board into your laptop's USB port.
6. In Arduino IDE:
   * Select **Tools -> Board -> ESP32 Arduino -> ESP32 Dev Module** (or ESP32S3).
   * Select **Tools -> Port -> COMx** (the COM port assigned to your board).
7. Click the **Upload** button (the arrow icon).
8. Once uploaded, open the **Serial Monitor** at **115200 baud**. You will see:
   `[CSI] Wi-Fi CSI Hardware Engine Activated!` and packets flowing!

---

## 4. How to Run the Detection System

### Option A: 1-Click Launch (Easiest)
Double-click `RUN_PROJECT.bat` in the root folder (or `Human_Detection_CSI\run_human_detector.bat`).

### Option B: From Terminal
```powershell
cd c:\Users\Akhil\OneDrive\Desktop\ruview\Human_Detection_CSI
python csi_detector_server.py --http-port 8888
```

Your browser will automatically open at:
👉 **http://localhost:8888**

---

## 5. What Happens When You Run It

* **Before you have the chip (Simulated Mode):**
  * The system automatically generates realistic RF multipath subcarrier data.
  * You can click **"Toggle Simulated Person"** on the dashboard to switch between `ROOM EMPTY` and `HUMAN DETECTED`.
  * Shows real-time respiration waveform (14–16 BPM) and subcarrier shifting.
* **When you plug in the ESP32 chip (Live Hardware Mode):**
  * The Python server auto-detects the ESP32 COM port.
  * The badge switches to `LIVE - ESP32 (COMx)`.
  * As you walk into the room or move in front of the Wi-Fi link, the 52 subcarrier bars fluctuate, variance spikes, and the status turns **GREEN (HUMAN DETECTED)**!
  * When you stay still, subtle chest movements modulate the 0.1–0.5 Hz band, displaying your **Breathing Rate (BPM)**!
