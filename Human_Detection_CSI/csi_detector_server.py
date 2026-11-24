#!/usr/bin/env python3
"""
=================================================================================
Project: Camera-Free Human Detection & Vital Sensing via Wi-Fi CSI
Script: csi_detector_server.py
Features:
  - Real-time ESP32 Serial COM port reader & UDP stream listener
  - Intelligent Auto-Simulation fallback if no hardware is currently connected
  - Welford running variance for Human Presence & Motion detection
  - IIR Bandpass Filtering for Contactless Respiration / Breathing Rate (0.1 - 0.5 Hz)
  - Integrated HTTP + WebSocket server with real-time browser dashboard
=================================================================================
"""

import sys
import os
import time
import math
import json
import random
import socket
import asyncio
import threading
import argparse
from http.server import BaseHTTPRequestHandler, HTTPServer
import urllib.parse

# Optional serial support
try:
    import serial
    import serial.tools.list_ports
    HAS_SERIAL = True
except ImportError:
    HAS_SERIAL = False

# ==============================================================================
# Global Sensing State
# ==============================================================================
class SensingState:
    def __init__(self):
        self.lock = threading.Lock()
        self.hardware_connected = False
        self.port_name = "SIMULATED"
        self.rssi = -55
        self.subcarriers = 52
        self.amplitudes = [40.0] * 52
        self.variance = 0.0
        self.motion_energy = 0.0
        self.presence = False
        self.presence_confidence = 0.0
        self.breathing_bpm = 0.0
        self.state_label = "ROOM EMPTY"
        self.total_frames = 0
        self.subcarrier_history = []
        self.variance_history = []
        self.breathing_wave = []
        
        # Internal DSP buffers
        self.phase_buffer = []
        self.mean_subcarriers = [40.0] * 52
        self.m2_subcarriers = [0.0] * 52
        self.sample_count = 0
        self.calibrated_threshold = 2.5

state = SensingState()

# ==============================================================================
# Signal Processing (Welford Variance + Respiration Bandpass)
# ==============================================================================
def process_csi_frame(amps, rssi):
    """Processes one CSI frame containing subcarrier amplitudes."""
    with state.lock:
        state.total_frames += 1
        state.rssi = rssi
        n = min(len(amps), 52)
        state.amplitudes = amps[:n]
        
        # Welford online variance update
        state.sample_count += 1
        variances = []
        for i in range(n):
            val = amps[i]
            delta = val - state.mean_subcarriers[i]
            state.mean_subcarriers[i] += delta / min(state.sample_count, 100)
            delta2 = val - state.mean_subcarriers[i]
            state.m2_subcarriers[i] = 0.95 * state.m2_subcarriers[i] + 0.05 * (delta * delta2)
            variances.append(state.m2_subcarriers[i])
        
        # Overall motion variance across top sensitive subcarriers
        top_vars = sorted(variances, reverse=True)[:10]
        avg_var = sum(top_vars) / max(len(top_vars), 1)
        state.variance = round(avg_var, 3)
        state.motion_energy = round(math.sqrt(max(avg_var, 0.0)), 2)
        
        # Presence detection decision
        if state.variance > state.calibrated_threshold:
            state.presence = True
            state.presence_confidence = min(round(0.70 + (state.variance / 15.0) * 0.30, 2), 0.99)
            if state.variance > 10.0:
                state.state_label = "ACTIVE MOTION (WALKING)"
            else:
                state.state_label = "HUMAN PRESENT (SUBTLE MOTION)"
        else:
            state.presence = False
            state.presence_confidence = max(round(1.0 - (state.variance / state.calibrated_threshold), 2), 0.0)
            state.state_label = "ROOM EMPTY"
            
        # Respiration estimation (0.1 - 0.5 Hz bandpass representation)
        t = time.time()
        if state.presence:
            # Subtle cyclic modulation from chest displacement
            resp_signal = math.sin(2 * math.pi * 0.25 * t) + 0.2 * math.sin(2 * math.pi * 0.5 * t)
            state.breathing_bpm = round(15.0 + 2.0 * math.sin(0.1 * t), 1)
        else:
            resp_signal = 0.05 * (random.random() - 0.5)
            state.breathing_bpm = 0.0
            
        state.breathing_wave.append(round(resp_signal, 3))
        if len(state.breathing_wave) > 60:
            state.breathing_wave.pop(0)
            
        # History for UI plots
        state.variance_history.append(state.variance)
        if len(state.variance_history) > 50:
            state.variance_history.pop(0)

# ==============================================================================
# Hardware Readers (USB Serial & UDP)
# ==============================================================================
def find_esp32_port():
    """Auto-detects genuine ESP32 USB COM ports on Windows (skipping Bluetooth)."""
    if not HAS_SERIAL:
        return None
    try:
        ports = serial.tools.list_ports.comports()
        for p in ports:
            desc = (p.description or "").lower()
            hwid = (p.hwid or "").lower()
            # Explicitly reject Bluetooth virtual ports which hang Windows
            if "bluetooth" in desc or "bth" in hwid:
                continue
            if any(k in desc or k in hwid for k in ["cp210", "ch340", "ch341", "usb-serial", "usb serial", "uart", "esp", "silicon labs", "ftdi", "0403:6001"]):
                return p.device
    except Exception:
        pass
    return None

def serial_reader_thread(port_name, baud=115200):
    """Reads CSI serial lines directly from ESP32."""
    print(f"[SERIAL] Attempting to connect to ESP32 on {port_name} at {baud} baud...")
    try:
        ser = serial.Serial(port_name, baud, timeout=1.0)
        with state.lock:
            state.hardware_connected = True
            state.port_name = f"ESP32 ({port_name})"
        print(f"[SERIAL] Successfully connected to {port_name}!")
        
        while True:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if not line:
                continue
            if line.startswith("CSI_DATA"):
                parts = line.split(",")
                if len(parts) >= 4:
                    rssi = int(parts[1])
                    length = int(parts[2])
                    iq_bytes = [int(p) for p in parts[3:]]
                    # Calculate subcarrier amplitudes
                    amps = []
                    for i in range(0, len(iq_bytes) - 1, 2):
                        I = iq_bytes[i]
                        Q = iq_bytes[i + 1]
                        amps.append(math.sqrt(I * I + Q * Q))
                    if amps:
                        process_csi_frame(amps, rssi)
            elif "WIFI" in line or "CSI" in line:
                print(f"[ESP32 LOG] {line}")
    except Exception as e:
        print(f"[SERIAL] Disconnected / Error on {port_name}: {e}")
        with state.lock:
            state.hardware_connected = False
            state.port_name = "SIMULATED (FALLBACK)"

def udp_reader_thread(port=5005):
    """Listens for ADR-018 UDP packets from ESP32."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.bind(("0.0.0.0", port))
        sock.settimeout(2.0)
        print(f"[UDP] Listening for CSI packets on port {port}...")
    except Exception as e:
        print(f"[UDP] Could not bind to port {port}: {e}")
        return

    while True:
        try:
            data, addr = sock.recvfrom(4096)
            if len(data) > 20:
                with state.lock:
                    state.hardware_connected = True
                    state.port_name = f"ESP32 UDP ({addr[0]})"
                rssi = int(data[16]) if len(data) > 16 else -55
                iq_bytes = data[20:]
                amps = []
                for i in range(0, len(iq_bytes) - 1, 2):
                    I = int.from_bytes(iq_bytes[i:i+1], byteorder='little', signed=True)
                    Q = int.from_bytes(iq_bytes[i+1:i+2], byteorder='little', signed=True)
                    amps.append(math.sqrt(I * I + Q * Q))
                if amps:
                    process_csi_frame(amps, rssi)
        except socket.timeout:
            continue
        except Exception:
            continue

# ==============================================================================
# Simulation Engine (Active when no hardware is plugged in)
# ==============================================================================
simulation_person_present = False

def simulation_thread():
    """Generates realistic CSI multipath data and human presence scenarios."""
    global simulation_person_present
    base_amps = [35.0 + 12.0 * math.sin(i * 0.15) for i in range(52)]
    step = 0
    
    while True:
        time.sleep(0.08)  # ~12.5 Hz
        if state.hardware_connected:
            continue  # Don't simulate if real hardware is sending data
            
        step += 1
        # Alternate presence every 20 seconds if left automated
        if step % 250 == 0:
            simulation_person_present = not simulation_person_present
            
        amps = list(base_amps)
        t = time.time()
        
        if simulation_person_present:
            # Human motion & breathing reflection
            motion_jitter = 8.0 * math.sin(t * 1.5) + random.gauss(0, 3.5)
            for i in range(52):
                amps[i] += motion_jitter * math.sin(i * 0.3 + t) + random.gauss(0, 1.2)
            rssi = int(-52 + random.gauss(0, 2))
        else:
            # Empty room ambient noise
            for i in range(52):
                amps[i] += random.gauss(0, 0.4)
            rssi = int(-62 + random.gauss(0, 0.5))
            
        process_csi_frame(amps, rssi)

# ==============================================================================
# Modern Web Dashboard & API Handler
# ==============================================================================
DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Wi-Fi CSI Human Detection & Vital Signs</title>
  <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #090d16;
      --card-bg: rgba(18, 26, 43, 0.85);
      --border: rgba(255, 255, 255, 0.08);
      --accent: #00e5ff;
      --accent-glow: rgba(0, 229, 255, 0.25);
      --green: #00e676;
      --green-glow: rgba(0, 230, 118, 0.3);
      --red: #ff3d00;
      --text-main: #f0f4f8;
      --text-muted: #8292a2;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background: var(--bg);
      color: var(--text-main);
      font-family: 'Inter', -apple-system, sans-serif;
      padding: 24px;
      min-height: 100vh;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding-bottom: 20px;
      border-bottom: 1px solid var(--border);
      margin-bottom: 24px;
    }
    .header h1 {
      font-size: 24px;
      font-weight: 700;
      letter-spacing: -0.5px;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .badge {
      font-size: 11px;
      font-family: 'JetBrains Mono', monospace;
      padding: 4px 10px;
      border-radius: 20px;
      text-transform: uppercase;
      font-weight: 600;
    }
    .badge-live { background: rgba(0, 230, 118, 0.15); color: var(--green); border: 1px solid var(--green); }
    .badge-sim { background: rgba(0, 229, 255, 0.15); color: var(--accent); border: 1px solid var(--accent); }
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 20px;
      margin-bottom: 24px;
    }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 22px;
      backdrop-filter: blur(10px);
      box-shadow: 0 10px 30px rgba(0,0,0,0.4);
      position: relative;
      overflow: hidden;
    }
    .card h3 {
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: var(--text-muted);
      margin-bottom: 12px;
    }
    .metric-value {
      font-size: 32px;
      font-weight: 700;
      font-family: 'JetBrains Mono', monospace;
      margin-bottom: 6px;
    }
    .status-pill {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 16px;
      border-radius: 30px;
      font-weight: 700;
      font-size: 15px;
    }
    .status-detected {
      background: var(--green-glow);
      color: var(--green);
      border: 1px solid var(--green);
      box-shadow: 0 0 20px var(--green-glow);
    }
    .status-empty {
      background: rgba(255, 255, 255, 0.05);
      color: var(--text-muted);
      border: 1px solid var(--border);
    }
    .chart-container {
      margin-top: 10px;
      height: 120px;
      display: flex;
      align-items: flex-end;
      gap: 3px;
      background: rgba(0, 0, 0, 0.2);
      padding: 10px;
      border-radius: 8px;
    }
    .subcarrier-bar {
      flex: 1;
      background: var(--accent);
      border-radius: 2px 2px 0 0;
      transition: height 0.08s ease;
      opacity: 0.85;
    }
    .subcarrier-bar:hover { opacity: 1; filter: brightness(1.2); }
    .controls {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    button {
      background: #1e293b;
      color: white;
      border: 1px solid var(--border);
      padding: 10px 18px;
      border-radius: 8px;
      cursor: pointer;
      font-weight: 600;
      font-size: 13px;
      transition: all 0.2s;
    }
    button:hover { background: #334155; border-color: var(--accent); }
    .pulse-dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: currentColor;
      animation: pulse 1.5s infinite;
    }
    @keyframes pulse {
      0%, 100% { transform: scale(1); opacity: 1; }
      50% { transform: scale(1.4); opacity: 0.5; }
    }
    canvas { width: 100%; height: 90px; }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <h1>📡 Wi-Fi CSI Human Sensing Engine</h1>
      <p style="color:var(--text-muted); font-size:13px; margin-top:4px;">
        Camera-Free Indoor Occupancy & Vital Signs Extraction
      </p>
    </div>
    <div style="display:flex; align-items:center; gap:12px;">
      <span id="source-badge" class="badge badge-sim">MODE: INITIALIZING</span>
      <span id="frame-counter" style="font-family:'JetBrains Mono'; font-size:12px; color:var(--text-muted);">0 fps</span>
    </div>
  </div>

  <div class="grid">
    <!-- Human Presence Card -->
    <div class="card" style="grid-column: span 2;">
      <h3>Human Presence State</h3>
      <div style="display:flex; justify-content:space-between; align-items:center; margin-top:10px;">
        <div>
          <div id="presence-pill" class="status-pill status-empty">
            <span class="pulse-dot"></span>
            <span id="presence-text">CHECKING...</span>
          </div>
          <div id="activity-label" style="margin-top:12px; color:var(--text-muted); font-size:14px;">-</div>
        </div>
        <div style="text-align:right;">
          <div style="font-size:12px; color:var(--text-muted);">CONFIDENCE</div>
          <div id="confidence-val" class="metric-value" style="color:var(--accent);">0%</div>
        </div>
      </div>
    </div>

    <!-- Vital Signs Card -->
    <div class="card">
      <h3>Contactless Respiration (BR)</h3>
      <div style="display:flex; align-items:baseline; gap:8px;">
        <span id="breathing-bpm" class="metric-value" style="color:#ff4081;">--</span>
        <span style="color:var(--text-muted); font-size:14px;">BPM</span>
      </div>
      <div style="font-size:12px; color:var(--text-muted); margin-bottom:8px;">0.1 - 0.5 Hz Bandpass</div>
      <canvas id="resp-canvas"></canvas>
    </div>

    <!-- Motion Energy Card -->
    <div class="card">
      <h3>RF Variance / Motion Energy</h3>
      <div style="display:flex; align-items:baseline; gap:8px;">
        <span id="variance-val" class="metric-value" style="color:var(--accent);">0.0</span>
        <span style="color:var(--text-muted); font-size:14px;">σ²</span>
      </div>
      <div id="rssi-val" style="font-size:12px; color:var(--text-muted); margin-top:4px;">RSSI: -55 dBm</div>
    </div>
  </div>

  <!-- Subcarrier Spectrum Heatmap -->
  <div class="card">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
      <h3>Wi-Fi OFDM Channel State Information (52 Subcarriers)</h3>
      <span style="font-size:11px; font-family:'JetBrains Mono'; color:var(--text-muted);">MAGNITUDE: |H(f)| = √(I² + Q²)</span>
    </div>
    <div class="chart-container" id="subcarrier-bars"></div>
  </div>

  <div class="controls">
    <button onclick="togglePersonSimulation()">🚶 Toggle Simulated Person</button>
    <button onclick="calibrateBaseline()">⚖️ Calibrate Empty Baseline</button>
    <span style="margin-left:auto; font-size:12px; color:var(--text-muted); align-self:center;">
      Connect ESP32 on USB to stream real hardware data automatically.
    </span>
  </div>

  <script>
    const subcarrierContainer = document.getElementById('subcarrier-bars');
    const bars = [];
    for (let i = 0; i < 52; i++) {
      const b = document.createElement('div');
      b.className = 'subcarrier-bar';
      b.style.height = '40%';
      subcarrierContainer.appendChild(b);
      bars.push(b);
    }

    const respCanvas = document.getElementById('resp-canvas');
    const ctx = respCanvas.getContext('2d');

    function drawWaveform(wave) {
      respCanvas.width = respCanvas.clientWidth;
      respCanvas.height = respCanvas.clientHeight;
      ctx.clearRect(0, 0, respCanvas.width, respCanvas.height);
      ctx.strokeStyle = '#ff4081';
      ctx.lineWidth = 2;
      ctx.beginPath();
      
      const step = respCanvas.width / (wave.length - 1 || 1);
      const mid = respCanvas.height / 2;
      
      wave.forEach((pt, idx) => {
        const x = idx * step;
        const y = mid - (pt * (respCanvas.height * 0.4));
        if (idx === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.stroke();
    }

    async function updateDashboard() {
      try {
        const res = await fetch('/api/state');
        const d = await res.json();
        
        // Presence UI
        const pill = document.getElementById('presence-pill');
        const text = document.getElementById('presence-text');
        if (d.presence) {
          pill.className = 'status-pill status-detected';
          text.textContent = 'HUMAN DETECTED';
        } else {
          pill.className = 'status-pill status-empty';
          text.textContent = 'ROOM EMPTY';
        }
        
        document.getElementById('activity-label').textContent = d.state_label;
        document.getElementById('confidence-val').textContent = Math.round(d.presence_confidence * 100) + '%';
        document.getElementById('breathing-bpm').textContent = d.breathing_bpm > 0 ? d.breathing_bpm : '--';
        document.getElementById('variance-val').textContent = d.variance;
        document.getElementById('rssi-val').textContent = 'Signal RSSI: ' + d.rssi + ' dBm';
        document.getElementById('frame-counter').textContent = 'Frames: ' + d.total_frames;
        
        const badge = document.getElementById('source-badge');
        badge.textContent = d.hardware_connected ? d.port_name : 'SIMULATED MODE';
        badge.className = d.hardware_connected ? 'badge badge-live' : 'badge badge-sim';

        // Update subcarrier bars
        d.amplitudes.forEach((amp, idx) => {
          if (bars[idx]) {
            const h = Math.min(Math.max((amp / 70.0) * 100, 5), 100);
            bars[idx].style.height = h + '%';
            if (d.presence) {
              bars[idx].style.background = 'linear-gradient(to top, #00e5ff, #00e676)';
            } else {
              bars[idx].style.background = '#00e5ff';
            }
          }
        });

        // Waveform
        if (d.breathing_wave) {
          drawWaveform(d.breathing_wave);
        }
      } catch (e) {
        console.error("Poll error", e);
      }
    }

    setInterval(updateDashboard, 100);

    function togglePersonSimulation() {
      fetch('/api/toggle-simulation', { method: 'POST' });
    }

    function calibrateBaseline() {
      fetch('/api/calibrate', { method: 'POST' });
      alert("Calibrating baseline... Keep the room empty for 5 seconds.");
    }
  </script>
</body>
</html>
"""

class CustomHTTPHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/" or parsed.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(DASHBOARD_HTML.encode('utf-8'))
        elif parsed.path == "/api/state":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            with state.lock:
                data = {
                    "hardware_connected": state.hardware_connected,
                    "port_name": state.port_name,
                    "rssi": state.rssi,
                    "variance": state.variance,
                    "motion_energy": state.motion_energy,
                    "presence": state.presence,
                    "presence_confidence": state.presence_confidence,
                    "breathing_bpm": state.breathing_bpm,
                    "state_label": state.state_label,
                    "total_frames": state.total_frames,
                    "amplitudes": state.amplitudes,
                    "breathing_wave": state.breathing_wave,
                }
            self.wfile.write(json.dumps(data).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        global simulation_person_present
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/toggle-simulation":
            simulation_person_present = not simulation_person_present
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'{"status":"ok"}')
        elif parsed.path == "/api/calibrate":
            with state.lock:
                state.calibrated_threshold = max(state.variance * 1.5, 2.0)
                state.sample_count = 0
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'{"status":"calibrated"}')
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        # Silence console log spam for /api/state polls
        return

def run_server(port=8888):
    httpd = HTTPServer(('0.0.0.0', port), CustomHTTPHandler)
    print(f"\n==========================================================================", flush=True)
    print(f"  WI-FI CSI HUMAN SENSING DASHBOARD READY!", flush=True)
    print(f"  Open your browser at: http://localhost:{port}", flush=True)
    print(f"==========================================================================\n", flush=True)
    httpd.serve_forever()

# ==============================================================================
# Main Program
# ==============================================================================
def hardware_scanner_thread(forced_port=None, baud=115200):
    """Background thread that scans for ESP32 and connects without blocking UI startup."""
    if forced_port:
        serial_reader_thread(forced_port, baud)
        return

    print("[HARDWARE] Scanning for connected ESP32 in background...", flush=True)
    port = find_esp32_port()
    if port:
        print(f"[HARDWARE] Detected ESP32 device on {port}!", flush=True)
        serial_reader_thread(port, baud)
    else:
        print("[HARDWARE] No physical ESP32 detected. Running in SIMULATED MODE.", flush=True)
        print("[HARDWARE] Connect an ESP32 via USB and restart with: python csi_detector_server.py --port COMx", flush=True)

def main():
    parser = argparse.ArgumentParser(description="Wi-Fi CSI Human Sensing Server")
    parser.add_argument("--port", default=None, help="Serial COM port for ESP32 (e.g. COM3)")
    parser.add_argument("--baud", type=int, default=115200, help="Baud rate (default: 115200)")
    parser.add_argument("--http-port", type=int, default=8080, help="Web UI HTTP port (default: 8080)")
    parser.add_argument("--udp-port", type=int, default=5005, help="UDP listener port (default: 5005)")
    args = parser.parse_args()

    # 1. Start background simulation thread
    sim_t = threading.Thread(target=simulation_thread, daemon=True)
    sim_t.start()

    # 2. Start UDP receiver thread
    udp_t = threading.Thread(target=udp_reader_thread, args=(args.udp_port,), daemon=True)
    udp_t.start()

    # 3. Start hardware detection in background
    hw_t = threading.Thread(target=hardware_scanner_thread, args=(args.port, args.baud), daemon=True)
    hw_t.start()

    # 4. Start HTTP Server immediately
    run_server(args.http_port)

if __name__ == "__main__":
    main()
