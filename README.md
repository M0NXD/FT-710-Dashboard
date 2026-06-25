# FT-710 CAT Control Dashboard

A compact, single-window desktop dashboard for the **Yaesu FT-710** transceiver.
It talks to the radio over CAT through [OmniRig](https://www.dxatlas.com/OmniRig/)
and gives you live metering, one-click band changes, quick power presets, and a
running TX log — all in a dark, always-readable panel.

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-GPLv3-green)

## Features

- **Live dashboard** — power (W), SWR, ALC, and a calibrated Yaesu-style
  S-meter, each on a colour-zoned bar with peak-hold on power and SWR.
- **Frequency & band readout** — current VFO A frequency (click to type a new
  one), VFO B, active band, and mode.
- **One-click band QSY** — 11 buttons (160 m → 6 m) jump to a sensible
  band-centre frequency and the conventional mode for that band.
- **Power presets** — separate, adjustable CW and SSB power levels with a
  one-click CW key/unkey and an SSB quick-set button.
- **VFO tools** — swap VFO A/B and toggle split with native OmniRig properties.
- **SWR alert** — visual warning plus an audible beep above a configurable
  threshold (1.0–3.0).
- **TX log** — the last 10 transmissions with frequency, band, mode, peak
  power, peak SWR, and duration.
- **Quality-of-life** — UTC clock, always-on-top pin, and a one-click OmniRig
  reconnect.

## Requirements

- **Windows** (uses the Win32 COM bridge and `winsound`)
- **Python 3.8+**
- **[OmniRig](https://www.dxatlas.com/OmniRig/)** installed and configured for
  your FT-710 on **Rig 1**
- Python packages: `pywin32`

Install the dependency:

```bash
pip install pywin32
```

## Setup

1. Install OmniRig and configure **Rig 1** for your FT-710 (correct COM port,
   baud rate, and the FT-710/FT-991A-style command set).
2. Make sure the radio's CAT settings match OmniRig (CAT baud rate, CAT RTS, and
   CAT timeout in the FT-710 menu).
3. Install the Python dependency (see above).

## Usage

Double-click `tune_up.pyw`, or run it from a terminal:

```bash
pythonw tune_up.pyw
```

The `.pyw` extension launches it without a console window. The app connects to
OmniRig Rig 1 on startup; if the connection fails, use the **Reconnect** button
after fixing OmniRig.

### Notes

- Click the frequency readout to type a new frequency (MHz, kHz, or Hz are all
  accepted).
- Meter readings rely on COM event callbacks from OmniRig. If events can't be
  connected, the dashboard shows a warning and meters stay blank, but tuning and
  control still work.
- The S-meter and SWR scales use calibration curves validated on an FT-710;
  treat them as good working approximations rather than lab-grade readings.

## Safety

This app can key the transmitter (CW key/unkey and power changes). Always
operate into a proper antenna or dummy load, observe the SWR alert, and follow
your local licensing and band-plan rules.

## License

Copyright (C) 2026 M0NXD.

Released under the [GNU General Public License v3.0](https://www.gnu.org/licenses/gpl-3.0.html).
This program is free software: you can redistribute it and/or modify it under
the terms of the GPL as published by the Free Software Foundation, either
version 3 of the License, or (at your option) any later version. It is
distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY.

## Acknowledgements

Built on [OmniRig](https://www.dxatlas.com/OmniRig/) by Alex Shovkoplyas, VE3NEA.
Yaesu and FT-710 are trademarks of Yaesu; this project is not affiliated with or
endorsed by Yaesu.
