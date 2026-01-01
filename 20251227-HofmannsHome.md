# KNX Projekt Bericht

## Projekt Statistiken
- **Gruppenadressen**: 1002
- **Geräte**: 25

## Gebäudestruktur
- 🏠 **Haus** (Gebäude)
  - 🪜 **EG** (Etage)
    - 🚪 **HWR** (Raum)
      - **Schaltschrank HWR** (Verteiler)
        - 🔌 **Enertex KNX IP Secure Router** (1.0.0)
        - 🔌 **IP-Systemgeräte - Zusatzfunktion** (1.0.9)
        - 🔌 **KNX Multi IO 570 (48I/O)** (1.0.40)
        - 🔌 **AKU-0816.03 Universalaktor 8-fach - Haus 01** (1.0.20) - AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC
        - 🔌 **Enertex KNX Dual PowerSupply 1280** (1.0.10)
        - 🔌 **Enertex KNX TP Secure Coupler** (1.1.0)
        - 🔌 **AKU-0816.03 Universalaktor 8-fach - Haus 03** (1.0.22) - AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC
        - 🔌 **AKU-0816.03 Universalaktor 8-fach - Haus 02** (1.0.21) - AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC
        - 🔌 **SCN-LOG1.02 Logikmodul Haus** (1.0.11) - SCN-LOG1.02 Logikmodul
    - 🚪 **Küche** (Raum)
    - 🚪 **Wohnzimmer** (Raum)
    - 🚪 **Esszimmer** (Raum)
    - 🚪 **Gästebad** (Raum)
  - 🪜 **OG** (Etage)
    - 🚪 **Arbeitszimmer** (Raum)
      - 🔌 **Home Assistant** (0.0.10) - Dummy
    - 🚪 **Schlafzimmer** (Raum)
    - 🚪 **Ankleide** (Raum)
    - 🚪 **Gästezimmer** (Raum)
    - 🚪 **Badezimmer** (Raum)
  - **Flur** (Flur)
- 🏠 **Gebäude Garage** (Gebäude)
  - 🚪 **Raum Garage** (Raum)
    - 🔌 **BE-GT20x.02 Glastaster II Smart** (1.1.50)
    - 🔌 **SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten** (1.1.101)
    - 🔌 **6131/31 Busch-Präsenzmelder Premium** (1.1.100)
    - 🔌 **Hörmann KNX-Gateway** (1.1.30) - KNX-Gateway
    - 🔌 **Raumtemperatursensor 1-fach UP** (1.1.201)
    - 🔌 **Garage Taster Tor-Nachbar** (1.1.51) - Taster Busankoppler AP WG 2fach Zweipunktbedienung
    - 🔌 **Garage Taster Tor-Feld** (1.1.52) - Taster Busankoppler AP WG 2fach Zweipunktbedienung
    - 🔌 **theLuxa P300 KNX - Garage Präsenz Aussen Tor** (1.1.102) - theLuxa P300 KNX
    - **Schaltschrank Garage** (Verteiler)
      - 🔌 **Enertex KNX PowerSupply 960³** (1.1.10)
      - 🔌 **KNX IO 410 - Binäreingang 4-fach Garage** (1.1.70) - KNX IO 410 (4I)
      - 🔌 **AKD-0424R2.02 - Dimmaktor Garage** (1.1.21) - AKD-0424R2.02 LED Controller 4 Kanal/RGBW, 2TE REG
      - 🔌 **AKU-0816.03 Universalaktor 8-fach - Garage** (1.1.20) - AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC
      - 🔌 **SCN-LOG1.02 Logikmodul Garage** (1.1.11) - SCN-LOG1.02 Logikmodul


## Geräte
> **Legende**:
> * **Flags**: `A`=Adresse, `P`=Programm, `Pa`=Parameter, `G`=Gruppen, `K`=Konfig. [X] = Programmiert, [.] = Nicht programmiert, [?] = Unbekannt

| Adresse | Name | Bestellnummer | Produkt Ref | Hersteller | Prog. Status |
|---|---|---|---|---|---|
| 0.0.10 | Home Assistant | Dummy | Dummy | Gira | `.....` |
| 1.0.0 | Enertex KNX IP Secure Router | 1164 | Enertex KNX IP Secure Router | Enertex Bayern GmbH | `APPaGK` |
| 1.0.9 | IP-Systemgeräte - Zusatzfunktion | 1164 - 1168 | IP-Systemgeräte - Zusatzfunktion | Enertex Bayern GmbH | `APPaGK` |
| 1.0.10 | Enertex KNX Dual PowerSupply 1280 | 1173 | Enertex KNX Dual PowerSupply 1280 | Enertex Bayern GmbH | `APPaGK` |
| 1.0.11 | SCN-LOG1.02 Logikmodul Haus | SCN-LOG1.02 | SCN-LOG1.02 Logikmodul | MDT Technologies | `APPaGK` |
| 1.0.20 | AKU-0816.03 Universalaktor 8-fach - Haus 01 | AKU-0816.03 | AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC | MDT Technologies | `APPaGK` |
| 1.0.21 | AKU-0816.03 Universalaktor 8-fach - Haus 02 | AKU-0816.03 | AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC | MDT Technologies | `APPaGK` |
| 1.0.22 | AKU-0816.03 Universalaktor 8-fach - Haus 03 | AKU-0816.03 | AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC | MDT Technologies | `APPaGK` |
| 1.0.40 | KNX Multi IO 570 (48I/O) | KNX Multi IO 570 | KNX Multi IO 570 (48I/O) | Weinzierl Engineering GmbH | `APPaGK` |
| 1.1.0 | Enertex KNX TP Secure Coupler | 1171 | Enertex KNX TP Secure Coupler | Enertex Bayern GmbH | `APPaGK` |
| 1.1.10 | Enertex KNX PowerSupply 960³ | 1152-3 | Enertex KNX PowerSupply 960³ | Enertex Bayern GmbH | `APPaGK` |
| 1.1.11 | SCN-LOG1.02 Logikmodul Garage | SCN-LOG1.02 | SCN-LOG1.02 Logikmodul | MDT Technologies | `APPaGK` |
| 1.1.20 | AKU-0816.03 Universalaktor 8-fach - Garage | AKU-0816.03 | AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC | MDT Technologies | `APPaGK` |
| 1.1.21 | AKD-0424R2.02 - Dimmaktor Garage | AKD-0424R2.02 | AKD-0424R2.02 LED Controller 4 Kanal/RGBW, 2TE REG | MDT Technologies | `APPaGK` |
| 1.1.30 | Hörmann KNX-Gateway | 4511630 | KNX-Gateway | Hörmann KG Verkaufsgesellschaft | `APPaGK` |
| 1.1.50 | BE-GT20x.02 Glastaster II Smart | BE-GT20x.02 | BE-GT20x.02 Glastaster II Smart | MDT Technologies | `APPaGK` |
| 1.1.51 | Garage Taster Tor-Nachbar | 5162 30 | Taster Busankoppler AP WG 2fach Zweipunktbedienung | Gira | `APPaGK` |
| 1.1.52 | Garage Taster Tor-Feld | 5162 30 | Taster Busankoppler AP WG 2fach Zweipunktbedienung | Gira | `APPaGK` |
| 1.1.70 | KNX IO 410 - Binäreingang 4-fach Garage | KNX IO 410 | KNX IO 410 (4I) | Weinzierl Engineering GmbH | `APPaGK` |
| 1.1.100 | 6131/31 Busch-Präsenzmelder Premium | 6131/31 | 6131/31 Busch-Präsenzmelder Premium | Busch-Jaeger | `APPaGK` |
| 1.1.101 | SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten | SCN-BWM55T.G2 | SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten | MDT Technologies | `APPaGK` |
| 1.1.102 | theLuxa P300 KNX - Garage Präsenz Aussen Tor | 1019610 | theLuxa P300 KNX | Theben AG | `APPaGK` |
| 1.1.201 | Raumtemperatursensor 1-fach UP | SCN-TS1UP.01 | Raumtemperatursensor 1-fach UP | MDT Technologies | `APPaGK` |
| 1.3.0 | KNX IO 534 CV (4D) | KNX IO 534 (4D) | KNX IO 534 CV (4D) | Weinzierl Engineering GmbH | `.....` |
| 1.3.0 | AKD-0230CC.02 LED Controller CC/CV 30 W / 230 V 2-Kanal | AKD-0230CC.02 | AKD-0230CC.02 LED Controller CC/CV 30 W / 230 V 2-Kanal | MDT Technologies | `.....` |

## Geräteparameter
> **Hinweis**: Zeigt explizit konfigurierte Parameter aus der Projektdatei.

### ⚙️ 1.0.0 - Enertex KNX IP Secure Router (Enertex KNX IP Secure Router)
| Parameter | Wert | Raw |
|---|---|---|
| BLOCKBROADCAST | **off** | `0` |
| BLOCKBROADCAST | **off** | `0` |
| PHYFILTER | **filter (default)** | `3` |
| PHYFILTER | **filter (default)** | `3` |
| LCGRPCONFIG_0_13 | **filter** | `3` |
| LCGRPCONFIG_0_13 | **filter** | `3` |

### ⚙️ 1.0.9 - IP-Systemgeräte - Zusatzfunktion (IP-Systemgeräte - Zusatzfunktion)
| Parameter | Wert | Raw |
|---|---|---|
| Status ext. time server valid | **on** | `1` |
| Report status of int. clock after restart | **on** | `1` |
| bSommerPolarität | **on** | `1` |
| Enable channel 1 to 10 | **off** | `0` |

### ⚙️ 1.0.10 - Enertex KNX Dual PowerSupply 1280 (Enertex KNX Dual PowerSupply 1280)
| Parameter | Wert | Raw |
|---|---|---|
| Enable timer functions | **Yes** | `1` |
| Enable calendar functions | **Yes** | `1` |
| Enable energy meter | **Yes** | `1` |
| Enable extreme values | **Yes** | `1` |
| Enable measured values | **Yes** | `1` |
| Enable time switches | **Yes** | `1` |
| Function | **Time receiver** | `1` |
| Request time after bus voltage recovery from bus | **Yes** | `1` |
| Value of  communication object "Send clock request" | **1** | `1` |
| Request date after bus voltage recovery from bus | **Yes** | `1` |
| Value of  communication object "Send date request" | **1** | `1` |
| Send workday status on change | **Yes** | `1` |
| City selection | **Berlin, 52.5°N, 13.4°E** | `11` |
| Does summer and winter time exist at the location | **Yes** | `1` |
| Send day/night status on change | **Yes** | `1` |
| Select data type for voltages | **Floating point - Volt [Dpt 14.027] ** | `1` |
| Send measured values cyclically | **Yes** | `1` |
| Send cyclically after | **900** | `900` |
| Send voltage Bus A cyclically | **Yes** | `1` |
| Send voltage Bus B cyclically | **Yes** | `1` |
| Send voltage Aux cyclically | **Yes** | `1` |
| Send current Bus A cyclically | **Yes** | `1` |
| Send current Bus B cyclically | **Yes** | `1` |
| Send current Aux cyclically | **Yes** | `1` |
| Send total current cyclically | **Yes** | `1` |
| Send power Bus A cyclically | **Yes** | `1` |
| Send power Bus B cyclically | **Yes** | `1` |
| Send power Aux cyclically | **Yes** | `1` |
| Send total power cyclically | **Yes** | `1` |
| Send temperature cyclically | **Yes** | `1` |
| Send telegram rate Bus A (current value) cyclically | **Yes** | `1` |
| Send telegram rate Bus A (average value since restart) cyclically | **Yes** | `1` |
| Send energy values cyclically | **Yes** | `1` |
| Send cyclically after | **900** | `900` |
| Send total emitted energy (lifetime) cyclically | **Yes** | `1` |
| Send total emitted energy (since power on) cyclically | **Yes** | `1` |
| Send total emitted energy (since analysis reset) cyclically | **Yes** | `1` |
| Send total consumed energy (lifetime) cyclically | **Yes** | `1` |
| Send total consumed energy (since switching on) cyclically | **Yes** | `1` |
| Send total consumed energy (since analysis reset) cyclically | **Yes** | `1` |
| Send measured values on change | **Yes** | `1` |
| Send total power on change | **Yes** | `1` |
| Send temperature on change | **Yes** | `1` |
| After restart, delay all telegrams to be sent by | **20** | `20` |

### ⚙️ 1.0.11 - SCN-LOG1.02 Logikmodul Haus (SCN-LOG1.02 Logikmodul)
| Parameter | Wert | Raw |
|---|---|---|
|     Town | **Berlin** | `1` |
| Objects for Date and Time | **separated objects** | `2` |
| Description of function | **Haus-Gruppe-Flur-Status** | `Haus-Gruppe-Flur-Status` |
| Output 2 sends | **if input On or Off** | `3` |
| Object for summer-/wintertime | **summer = value 1** | `1` |
| Object for sunrise | **send value 1** | `1` |
| Object for dusk | **send value 1** | `1` |
| Object for sundown | **send value 1** | `1` |
| Main function | **Logoc gate / Inverter** | `2` |
| Additional text | **Status Rückmeldungen Licht im Flur** | `Status Rückmeldungen Licht im Flur` |
|     Value if input ON | **value 0** | `0` |
|     Send with failure | **value 1** | `1` |
|     Select internal object | **254** | `254` |
| Additional text | **Status Rückmeldungen Licht im Badezimmer** | `Status Rückmeldungen Licht im Badezimmer` |
| Description of function | **Haus-Gruppe-Badezimmer** | `Haus-Gruppe-Badezimmer` |
| Send Operating cyclic | **1 h** | `60` |
| Output 2 | **active** | `1` |
| Output 3 | **active** | `1` |
| input 3 | **active** | `1` |
| Input 4 | **active** | `1` |
| Calculation of output | **average value** | `2` |
| Requet inputs after reset | **active** | `255` |
| Input 2 | **active** | `1` |
| Data point type | **DPT 5.001  1Byte percent value (0...100%)  ** | `42` |
| Input 3 | **active** | `1` |
| Input 2 | **active** | `1` |
| Input 3 | **active** | `1` |
| Startup delaytime | **60** | `60` |
| Requet inputs after reset | **active** | `255` |
| The output sends only if all inputs are valid | **active** | `1` |
| Requet inputs after reset | **active** | `255` |
|     Value | **value 1** | `1` |
| Condition 2: AND | **active** | `1` |
| Text for condition 1 | **Spots** | `Spots` |
| Text for condition 2 | **Treppenlicht** | `Treppenlicht` |
| Condition 3: AND | **active** | `1` |
| The output sends only if all inputs are valid | **active** | `1` |
| Description of function | **Haus-Gruppe-Küche-Status** | `Haus-Gruppe-Küche-Status` |
| Additional text | **Status Rückmeldungen Licht in der Lüche** | `Status Rückmeldungen Licht in der Lüche` |
| Main function | **Logoc gate / Inverter** | `2` |
| Input 2 | **active** | `1` |
| Main function | **Logoc gate / Inverter** | `2` |
| Requet inputs after reset | **active** | `255` |
| The output sends only if all inputs are valid | **active** | `1` |

### ⚙️ 1.0.20 - AKU-0816.03 Universalaktor 8-fach - Haus 01 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Parameter | Wert | Raw |
|---|---|---|
| Channel A / B | **Switch, Staircase, Heating control** | `1` |
| Channel C / D | **Switch, Staircase, Heating control** | `1` |
| Channel E / F | **Switch, Staircase, Heating control** | `1` |
| Channel G / H | **Switch, Staircase, Heating control** | `1` |

### ⚙️ 1.0.21 - AKU-0816.03 Universalaktor 8-fach - Haus 02 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Parameter | Wert | Raw |
|---|---|---|
| Device startup time | **3** | `3` |
| Channel A / B | **Switch, Staircase, Heating control** | `1` |
| Channel C / D | **Switch, Staircase, Heating control** | `1` |
| Channel E / F | **Switch, Staircase, Heating control** | `1` |
| Channel G / H | **Switch, Staircase, Heating control** | `1` |

### ⚙️ 1.0.22 - AKU-0816.03 Universalaktor 8-fach - Haus 03 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Parameter | Wert | Raw |
|---|---|---|
| Device startup time | **4** | `4` |
| Channel A / B | **Switch, Staircase, Heating control** | `1` |
| Channel C / D | **Switch, Staircase, Heating control** | `1` |
| Channel E / F | **Switch, Staircase, Heating control** | `1` |
| Channel G / H | **Switch, Staircase, Heating control** | `1` |

### ⚙️ 1.0.40 - KNX Multi IO 570 (48I/O) (KNX Multi IO 570 (48I/O))
| Parameter | Wert | Raw |
|---|---|---|
| Timer type | **Impulse (Staircase)** | `3` |
| User control | **Short / Long** | `1` |
| Function name | **Küche** | `Küche` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Send value | **Shutter** | `6` |
| Heartbeat | **Enabled ** | `1` |
|     Cycle time | **1 h** | `60` |
| Display synchronization | **Enabled ** | `1` |
| Channel function 1 | **Binary Input** | `3` |
| Channel function 2 | **Binary Input** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **EG-Flur-Deckenleuchte** | `EG-Flur-Deckenleuchte` |
| Name | **EG-Flur-Spots** | `EG-Flur-Spots` |
| Channel function 3 | **Binary Input** | `3` |
| Channel function 4 | **Binary Input** | `3` |
| Channel function 5 | **Binary Input** | `3` |
| Channel function 6 | **Binary Input** | `3` |
| Channel function 7 | **Binary Input** | `3` |
| Channel function 8 | **Binary Input** | `3` |
| Channel function 9 | **Binary Input** | `3` |
| Channel function 10 | **Binary Input** | `3` |
| Channel function 11 | **Binary Input** | `3` |
| Channel function 12 | **Binary Input** | `3` |
| Channel function 13 | **Binary Input** | `3` |
| Channel function 14 | **Binary Input** | `3` |
| Channel function 15 | **Binary Input** | `3` |
| Channel function 16 | **Binary Input** | `3` |
| Channel function 17 | **Binary Input** | `3` |
| Channel function 18 | **Binary Input** | `3` |
| Channel function 19 | **Binary Input** | `3` |
| Channel function 20 | **Binary Input** | `3` |
| Channel function 21 | **Binary Input** | `3` |
| Channel function 22 | **Binary Input** | `3` |
| Channel function 23 | **Binary Input** | `3` |
| Channel function 24 | **Binary Input** | `3` |
| Channel function 25 | **Binary Input** | `3` |
| Channel function 26 | **Binary Input** | `3` |
| Channel function 27 | **Binary Input** | `3` |
| Channel function 28 | **Binary Input** | `3` |
| Name | **Flur-Treppenlicht** | `Flur-Treppenlicht` |
| Name | **OG-Flur-Deckenleuchte** | `OG-Flur-Deckenleuchte` |
| Name | **OG-Flur-Wandleuchte** | `OG-Flur-Wandleuchte` |
| Name | **EG-Gaestebad-Licht** | `EG-Gaestebad-Licht` |
| Name | **EG-HWR-Deckenleuchte** | `EG-HWR-Deckenleuchte` |
| Name | **EG-Kueche-Hinten-Licht** | `EG-Kueche-Hinten-Licht` |
| Name | **EG-Kueche-Vorn-Licht** | `EG-Kueche-Vorn-Licht` |
| Name | **EG-Esszimmer-Deckenleuchte** | `EG-Esszimmer-Deckenleuchte` |
| Name | **EG-WZ-Couchtisch** | `EG-WZ-Couchtisch` |
| Name | **EG-WZ-Wandleuchten** | `EG-WZ-Wandleuchten` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **EG-WZ-Steckdosen** | `EG-WZ-Steckdosen` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Gaestezimmer-Deckenleuchte** | `OG-Gaestezimmer-Deckenleuchte` |
| Function of output a on press | **Toggle** | `3` |
| Function of output a on release | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Name | **OG-Badezimmer-Spots** | `OG-Badezimmer-Spots` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Badezimmer-Deckenleuchte** | `OG-Badezimmer-Deckenleuchte` |
| Function of output a on press | **Toggle** | `3` |
| Function of output a on release | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Name | **OG-Badezimmer-Spiegel** | `OG-Badezimmer-Spiegel` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Arbeitszimmer-Deckenleuchte** | `OG-Arbeitszimmer-Deckenleuchte` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Schlafzimmer-Deckenleuchte** | `OG-Schlafzimmer-Deckenleuchte` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Schlafzimmer-Ankleide** | `OG-Schlafzimmer-Ankleide` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **Panik-Taster** | `Panik-Taster` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Switch on** | `1` |
| Name | **Aussenbeleuchtung** | `Aussenbeleuchtung` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **Taster-Zentral-Aus** | `Taster-Zentral-Aus` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Switch on** | `1` |
| Function of output a on long press | **Switch off** | `2` |
| Name | **Lichtszene-Kueche** | `Lichtszene-Kueche` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **Lichtszene-Badezimmer** | `Lichtszene-Badezimmer` |
| Name | **Lichtszene-Flur** | `Lichtszene-Flur` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Name | **OG-Arbeitszimmer-Schreibtisch** | `OG-Arbeitszimmer-Schreibtisch` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| User control | **Short / Long** | `1` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on long press | **Toggle** | `3` |
| Function name | **Flur 01** | `Flur 01` |
| Function name | **Flur 02** | `Flur 02` |

### ⚙️ 1.1.0 - Enertex KNX TP Secure Coupler (Enertex KNX TP Secure Coupler)
| Parameter | Wert | Raw |
|---|---|---|
| PHYFILTER | **filter (default)** | `3` |
| PHYFILTER | **filter (default)** | `3` |
| BLOCKBROADCAST | **off** | `0` |
| BLOCKBROADCAST | **off** | `0` |
| Enertex special functions | **on** | `1` |
| LCCONFIG ACK_PYH | **on** | `0` |
| LCCONFIG PYH_REPEAT | **on** | `1` |
| telegram rate Subline | **45** | `45` |
| LCCONFIG ACK_PYH | **on** | `0` |
| LCCONFIG PYH_REPEAT | **on** | `1` |
|  | **45** | `45` |
| COUPL_SERV_CONTROL | **on** | `1` |
| LCGRPCONFIG_14_15 | **filter** | `3` |
| LCGRPCONFIG_16_31 | **filter** | `3` |
| LCGRPCONFIG_14_15 | **filter** | `3` |
| LCGRPCONFIG_16_31 | **filter** | `3` |

### ⚙️ 1.1.10 - Enertex KNX PowerSupply 960³ (Enertex KNX PowerSupply 960³)
| Parameter | Wert | Raw |
|---|---|---|
| Variant | **Enertex KNX PowerSupply 960³** | `1` |
| Enable timer functions | **Yes** | `1` |
| Enable calendar functions | **Yes** | `1` |
| Enable time switches | **Yes** | `1` |
| Enable energy meter | **Yes** | `1` |
| Enable extreme values | **Yes** | `1` |
| Enable measured values | **Yes** | `1` |
| Automatic changeover between winter and summer time | **Yes** | `1` |
| Send daylight saving time status on change | **Yes** | `1` |
| Send workday status on change | **Yes** | `1` |
| City selection | **Berlin, 52.5°N, 13.4°E** | `11` |
| Does summer and winter time exist at the location | **Yes** | `1` |
| Send day/night status on change | **Yes** | `1` |
| Select data type for voltages | **Floating point - Volt [Dpt 14.027] ** | `1` |
| Function | **Time receiver** | `1` |
| Request time after bus voltage recovery from bus | **Yes** | `1` |
| Request date after bus voltage recovery from bus | **Yes** | `1` |
| Value of  communication object "Send date request" | **1** | `1` |
| Value of  communication object "Send clock request" | **1** | `1` |
| Number of timers | **3** | `3` |
| Designation for CO | **Alles Aus 9:00 Uhr** | `Alles Aus 9:00 Uhr` |
| Telegram 1 | **Switching** | `1` |
| Value specification | **On** | `1` |
| Designation for time switch | **Alles Aus** | `Alles Aus` |
| Telegram 1 | **Switching** | `1` |
| Value specification | **Raise - 100%** | `9` |
| Designation for CO | **Dämmerung-Dimm-100%** | `Dämmerung-Dimm-100%` |
| Designation for time switch | **Aussen Abends** | `Aussen Abends` |
| Hour | **9** | `9` |
| Offset for astro time (minutes) | **30** | `30` |
| Designation for the switching time | **9:00Uhr** | `9:00Uhr` |
| Designation for CO | **Dämmerung Ein** | `Dämmerung Ein` |
| Designation for the switching time | **Dämmerung** | `Dämmerung` |
|  | **Astro - Sunset** | `2` |
| Offset for astro time (minutes) | **10** | `10` |
| Designation for CO | **Morgens Ein** | `Morgens Ein` |
| The following preconfigured telegrams are sent at the switching time | **2** | `1` |
| Number of switching times | **2** | `2` |
| Send time periods status on change | **Yes** | `1` |
|  | **Simple date** | `0` |
| Telegram 2 | **Switching** | `1` |
| Designation for CO | **Mitternacht Aus** | `Mitternacht Aus` |
| Hour | **23** | `23` |
| Minute | **59** | `59` |
| Designation for the switching time | **Mitternacht** | `Mitternacht` |
| Prefixed text for status output | **Yes** | `1` |
| Send measured values cyclically | **Yes** | `1` |
| Send cyclically after | **900** | `900` |
| Send voltage Bus cyclically | **Yes** | `1` |
| Send voltage Aux A cyclically | **Yes** | `1` |
| Send voltage Aux B cyclically | **Yes** | `1` |
| Send current Bus cyclically | **Yes** | `1` |
| Send current Aux A cyclically | **Yes** | `1` |
| Send current Aux B cyclically | **Yes** | `1` |
| Send total current cyclically | **Yes** | `1` |
| Send power Bus cyclically | **Yes** | `1` |
| Send power Aux A cyclically | **Yes** | `1` |
| Send power Aux B cyclically | **Yes** | `1` |
| Send total power cyclically | **Yes** | `1` |
| Send temperature cyclically | **Yes** | `1` |
| Send telegram rate Bus (current value) cyclically | **Yes** | `1` |
| Send telegram rate Bus (average value since restart) cyclically | **Yes** | `1` |
| Send energy values cyclically | **Yes** | `1` |
| Send total emitted energy (lifetime) cyclically | **Yes** | `1` |
| Send total emitted energy (since power on) cyclically | **Yes** | `1` |
| Send total emitted energy (since analysis reset) cyclically | **Yes** | `1` |
| Send total consumed energy (lifetime) cyclically | **Yes** | `1` |
| Send total consumed energy (since switching on) cyclically | **Yes** | `1` |
| Send total consumed energy (since analysis reset) cyclically | **Yes** | `1` |
| Send cyclically after | **900** | `900` |
| Send measured values on change | **Yes** | `1` |
| Send total power on change | **Yes** | `1` |
| Send temperature on change | **Yes** | `1` |
| Designation for switching time change hour-CO | **Mitternacht** | `Mitternacht` |
| Designation for switching time change minute-CO | **Mitternacht** | `Mitternacht` |
| Designation for switching time change hour-CO | **Dämmerung** | `Dämmerung` |
| Designation for switching time change minute-CO | **Dämmerung** | `Dämmerung` |
| Designation for "change switching time hour"-CO | **Alles Aus** | `Alles Aus` |
| Designation for "change switching time minute"-CO | **Alles Aus** | `Alles Aus` |
| After restart, delay all telegrams to be sent by | **30** | `30` |
| Designation for time switch | **Aussen Morgens** | `Aussen Morgens` |
| Number of switching times | **2** | `2` |
| Telegram 1 | **Switching** | `1` |
| Value specification | **On** | `1` |
| Telegram 2 | **Switching** | `1` |
| Designation for CO | **Morgens Aus** | `Morgens Aus` |
| Designation for the switching time | **Morgens Ein** | `Morgens Ein` |
| Designation for switching time change hour-CO | **Morgend Ein Stunde** | `Morgend Ein Stunde` |
| Hour | **6** | `6` |
| Minute | **25** | `25` |
| Designation for switching time change minute-CO | **Morgens Ein Minute** | `Morgens Ein Minute` |
| Designation for the switching time | **Morgens Aus** | `Morgens Aus` |
| Designation for switching time change hour-CO | **Morgens Aus Stunde** | `Morgens Aus Stunde` |
| Designation for switching time change minute-CO | **Morgens Aus Minute** | `Morgens Aus Minute` |
|  | **Astro - Sunrise** | `1` |
| The following preconfigured telegrams are sent at the switching time | **2** | `1` |
| Offset for astro time (minutes) | **-10** | `-10` |
|  | **Period 1** | `1` |
| Number of time periods | **1** | `1` |
| Name of the time period | **Oktober bis Februar** | `Oktober bis Februar` |
| Name of the time period active-CO | **Oktober bis Februar** | `Oktober bis Februar` |
| month | **10** | `10` |
| Number of days | **155** | `155` |

### ⚙️ 1.1.11 - SCN-LOG1.02 Logikmodul Garage (SCN-LOG1.02 Logikmodul)
| Parameter | Wert | Raw |
|---|---|---|
|     Town | **Berlin** | `1` |
| Objects for Date and Time | **separated objects** | `2` |
| Main function | **Telegram multipler/sequencer** | `9` |
| Output 2 sends | **if input On or Off** | `3` |
| Object for summer-/wintertime | **summer = value 1** | `1` |
| Object for sunrise | **send value 1** | `1` |
| Object for dusk | **send value 1** | `1` |
| Object for sundown | **send value 1** | `1` |
| Description of function | **Garage - Alles Auf/Zu** | `Garage - Alles Auf/Zu` |
| Additional text | **Steuere Garagentor & Gartentor** | `Steuere Garagentor & Gartentor` |
|     Value if input ON | **value 0** | `0` |
|     Send with failure | **value 1** | `1` |
|     Select internal object | **254** | `254` |
| Logic function | **OR** | `2` |
| Send condition | **at input telegram** | `2` |
| Main function | **Min/Max/Average value** | `14` |
| Output 2 | **active** | `1` |
| Output 3 | **active** | `1` |
| input 3 | **active** | `1` |
| Input 4 | **active** | `1` |
| Calculation of output | **average value** | `2` |
| Requet inputs after reset | **active** | `255` |
| Description of function | **Garage.Innen.LED-Band.Dimmwert** | `Garage.Innen.LED-Band.Dimmwert` |
| Data point type | **DPT 5.001  1Byte percent value (0...100%)  ** | `42` |
| Input 2 | **active** | `1` |
| Input 3 | **active** | `1` |
| Startup delaytime | **60** | `60` |
| Requet inputs after reset | **active** | `255` |
| The output sends only if all inputs are valid | **active** | `1` |

### ⚙️ 1.1.20 - AKU-0816.03 Universalaktor 8-fach - Garage (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Parameter | Wert | Raw |
|---|---|---|
| Channel A / B | **Blinds, Shutter** | `0` |
| Channel C / D | **Switch, Staircase, Heating control** | `1` |
|     Function Channel B | **not active** | `0` |
| Channel E / F | **Switch, Staircase, Heating control** | `1` |
|     Function Channel F | **not active** | `0` |
| Device startup time | **3** | `3` |

### ⚙️ 1.1.21 - AKD-0424R2.02 - Dimmaktor Garage (AKD-0424R2.02 LED Controller 4 Kanal/RGBW, 2TE REG)
| Parameter | Wert | Raw |
|---|---|---|
|     Toggle light Day/Night | **direct** | `1` |
|     Value for Day/Night | **Day = 0 / Night = 1** | `0` |
| Startup delay | **5** | `5` |
| Central objects | **active** | `1` |
|     Scenes | **not active** | `0` |
| Switching on behaviour | **time dependent dimming (from R5.0)** | `2` |
| Switching on behaviour | **time dependent dimming (from R5.0)** | `2` |
| Switching on behaviour | **time dependent dimming** | `2` |
| Description of objects | **Garage-Innen-LED-Band-Tuer** | `Garage-Innen-LED-Band-Tuer` |
| Behaviour after reset | **deactivation** | `0` |
| Device selection | **AKD-0424R2.02 (4x2A, without relay contact)** | `1` |
| Description of objects | **Garage-Innen-LED-Band-Nachbar** | `Garage-Innen-LED-Band-Nachbar` |
| Switch chanel off with relative dimming | **active** | `1` |
|     Town | **Berlin** | `1` |
| Central objects | **active** | `1` |
| Block object 1 - datapoint type | **1Bit Object** | `1` |
| Switch chanel off with relative dimming | **active** | `1` |
| Behaviour after reset | **deactivation** | `0` |
| Central objects | **active** | `1` |
|     Scenes | **not active** | `0` |
| Description of objects | **Garage-Innen-LED-Band-Feld** | `Garage-Innen-LED-Band-Feld` |
| Switch chanel off with relative dimming | **active** | `1` |
| Behaviour after reset | **deactivation** | `0` |
| Central objects | **active** | `1` |
|     Scenes | **not active** | `0` |
| Description of objects | **Garage-Innen-LED-Band-Mitte** | `Garage-Innen-LED-Band-Mitte` |
| Behaviour after reset | **deactivation** | `0` |
| Central objects | **active** | `1` |
|     Scenes | **not active** | `0` |
| Send status of dimming value | **at change 20%** | `51` |
| Send status objects cyclic | **1 h** | `3600` |
| Send status of dimming value | **at change 20%** | `51` |
| Send status objects cyclic | **1 h** | `3600` |
| Send status of dimming value | **at change 20%** | `51` |
| Send status objects cyclic | **1 h** | `3600` |
| Send status of dimming value | **at change 20%** | `51` |
| Send status objects cyclic | **1 h** | `3600` |
|     Relative dimming Brightness (V) | **2** | `2` |
|     Absolute dimming | **2** | `2` |
| Switch chanel off with relative dimming | **active** | `1` |
| Send "Operation" cyclic | **1 h** | `60` |

### ⚙️ 1.1.30 - Hörmann KNX-Gateway (KNX-Gateway)
| Parameter | Wert | Raw |
|---|---|---|
| Antriebstyp | **DC Antrieb (SupraMatic E)** | `8` |
| Verhalten nach Busspannungswiederkehr | **Aktuellen Zustand senden** | `1` |
| Bezeichnung des Antriebs | **Garagentor** | `Garagentor` |

### ⚙️ 1.1.50 - BE-GT20x.02 Glastaster II Smart (BE-GT20x.02 Glastaster II Smart)
| Parameter | Wert | Raw |
|---|---|---|
|     Labelling "Value" | **Garagentor** | `Garagentor` |
| Status value 3 | **14 Byte DPT 16.000 String** | `24` |
|  | **9-12 LED** | `3` |
| Number of lines | **2** | `1` |
| LED behaviour in Standby | **Functions LEDs** | `2` |
|         Status element right | **Status value 2** | `3` |
| Status value 1 | **2 Byte DPT 9.001 Temperature [°C]** | `6` |
| Status value 2 | **2 Byte DPT 9.007 Humidity [%]** | `8` |
| Standby display during "Day" | **Standby over full screen** | `2` |
| Number of lines | **2** | `1` |
| Function 5/6 | **two-button function** | `1` |
| Slap / Cleaning function | **active** | `1` |
| Slap function for short keypress | **toggle** | `2` |
| Display switch-off in Standby | **active for "Day" and "Night"** | `3` |
| Value Day / Night | **Night = 1 / Day = 0** | `2` |
| Brightness | **brightness level 10** | `9` |
| Behaviour of LEDs on bus power return | **request LED objects** | `1` |
|  | **not active** | `0` |
| Behaviour on bus power return | **request external logic objects** | `1` |
| Function 7/8 | **two-button function** | `1` |
| 2. level / 12 functions | **active** | `1` |
|     Font size: third status line | **small** | `1` |
|     Level 2, lower buttons | **not active** | `255` |
|         Status element | **Status text 1 (via object 147)** | `6` |
| 2. level / 8 functions | **active** | `1` |
| Description of function | **Garage-Innen-LED-Band-A-B** | `Garage-Innen-LED-Band-A-B` |
| Logic object A (external) | **normally active, with preallocation "0"** | `1` |
| Logic object B (external) | **normally active, with preallocation "0"** | `1` |
| Minimum brightness "Night" | **5%** | `4` |
| Description of function | **Garage-Innen-LED-Band-C-D** | `Garage-Innen-LED-Band-C-D` |
| Minimum brightness "Day" | **20%** | `6` |
| Description of function | **Garage-Innen-LED-Band-Gesamt** | `Garage-Innen-LED-Band-Gesamt` |
| Logic object A (external) | **normally active, with preallocation "0"** | `1` |
| Logic object B (external) | **normally active, with preallocation "0"** | `1` |
| Logic object A (external) | **normally active, with preallocation "0"** | `1` |
| Logic object B (external) | **normally active, with preallocation "0"** | `1` |
| Setting Logic 4 | **OR** | `0` |
| Description of function | **Garage-Aussen-Downlights** | `Garage-Aussen-Downlights` |
| Logic object A (external) | **normally active, with preallocation "0"** | `1` |
| Logic object B (external) | **normally active, with preallocation "0"** | `1` |
| Function 3/4 | **single-button function** | `2` |
| Function 1/2 | **single-button function** | `2` |
| Time for long keypress | **5,5 s** | `38268` |
|     for "Day" (value ON) | **light blue (pastel)** | `137` |
| LED reacts to: | **external object and button activation** | `9` |
| LED reacts to: | **external object and button activation** | `9` |
|     for "Day" (value ON) | **red** | `240` |
|     for "Night" (value ON) | **red** | `240` |
| LED reacts to: | **external object and button activation** | `9` |
| LED reacts to: | **external object and button activation** | `9` |
| LED reacts to: | **external object and button activation** | `9` |
| LED reacts to: | **external object and button activation** | `9` |
|     for "Night" (value ON) | **light blue (pastel)** | `137` |
|     Level 2, middle buttons | **not active** | `255` |
|     for "Day" (value ON) | **sun orange** | `241` |
|  | **active** | `1` |
|     for "Night" (value ON) | **sun orange** | `241` |
|  | **active** | `1` |
|     for "Day" (value ON) | **sun orange** | `241` |
|     for "Night" (value ON) | **sun orange** | `241` |
|  | **active** | `1` |
| LED reacts to: | **external object and button activation** | `9` |
| LED reacts to: | **external object and button activation** | `9` |
|     for "Night" (value ON) | **red** | `240` |
|     for "Day" (value ON) | **red** | `240` |
|     for "Day" (value ON) | **light blue (pastel)** | `137` |
|     for "Night" (value ON) | **light blue (pastel)** | `137` |
|     for "Day" (value ON) | **light blue (pastel)** | `137` |
|     for "Night" (value ON) | **light blue (pastel)** | `137` |
|     Output inverted | **active** | `1` |
|     Object type | **1 Byte DPT 5.005 Decimal value (0...255)** | `3` |
|     Behaviour for "Day" (value ON) | **flashing** | `1` |
|     Behaviour for "Night" (value ON) | **flashing** | `1` |
|     for "Day" (value ON) | **light blue (pastel)** | `137` |
|     for "Night" (value ON) | **light blue (pastel)** | `137` |
|     Behaviour for "Night" (value ON) | **flashing** | `1` |
|     Behaviour for "Day" (value ON) | **flashing** | `1` |
|  | **active** | `1` |
|  | **active** | `1` |
|  | **active** | `1` |
|  | **active** | `1` |
|  | **active** | `1` |
|  | **active** | `1` |
| Startup time | **10** | `10` |

### ⚙️ 1.1.51 - Garage Taster Tor-Nachbar (Taster Busankoppler AP WG 2fach Zweipunktbedienung)
| Parameter | Wert | Raw |
|---|---|---|
| Wippe 2 rechts
Tasten 3 (oben) und 4 (unten) | **Tastenfunktion** | `0` |
| Funktion | **Jalousie** | `3` |
| Befehl beim Drücken der
Taste | **EIN** | `1` |

### ⚙️ 1.1.52 - Garage Taster Tor-Feld (Taster Busankoppler AP WG 2fach Zweipunktbedienung)
| Parameter | Wert | Raw |
|---|---|---|
| Wippe 2 rechts
Tasten 3 (oben) und 4 (unten) | **Tastenfunktion** | `0` |
| Funktion | **Jalousie** | `3` |
| Befehl beim Drücken der
Taste | **EIN** | `1` |

### ⚙️ 1.1.70 - KNX IO 410 - Binäreingang 4-fach Garage (KNX IO 410 (4I))
| Parameter | Wert | Raw |
|---|---|---|
| Function of output a on release | **Switch off** | `2` |
| Function of output a on press | **Switch on** | `1` |
| LED visualization bottom | **Disabled** | `0` |
| Name | **** | `` |
| Name | **Garage.Fenster** | `Garage.Fenster` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output b on short press | **Toggle** | `3` |
| Dimming function | **Off / Dim darker** | `1` |
| Shutter function | **Toggle direction** | `2` |
| Dimming function | **Toggle direction** | `2` |
| Shutter function | **Toggle direction** | `2` |
| Name | **** | `` |
| User control | **Short = Drive / Short = Step-Stop** | `1` |
| Function of output a on press | **Toggle** | `3` |
| Name | **** | `` |
| Function of output a on short press | **Toggle** | `3` |
| Function of output a on press | **Switch on** | `1` |
| Timer type | **Switch-on and -off delay** | `2` |
| Function of output a on long press | **Toggle** | `3` |
| Heartbeat | **Enabled ** | `1` |
| Device name | **Binäreingang Garage** | `Binäreingang Garage` |
|     Cycle time | **1 h** | `60` |
| Function of output a on short press | **Switch on** | `1` |
| Function | **Switching** | `1` |
| User control | **Short = Drive / Short = Step-Stop** | `1` |

### ⚙️ 1.1.100 - 6131/31 Busch-Präsenzmelder Premium (6131/31 Busch-Präsenzmelder Premium)
| Parameter | Wert | Raw |
|---|---|---|
| Applikation | **Melder** | `1` |
| Erweiterte Parameter einblenden | **ja** | `1` |
| Erweiterte Parameter einblenden | **ja** | `1` |
| Channel function table | **45066** | `45066` |
| Helligkeitsdifferenz für sofortiges Senden (%) | **60** | `60` |
| Helligkeitsdifferenz für sofortiges Senden (%) | **60** | `60` |
| Eingang Slave | **ja** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Erweiterte Parameter einblenden | **ja** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Objekt Eingang externer Taster nutzen | **ja** | `1` |
| Objekt Eingang externer Taster nutzen | **ja** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Wert für reduzierte Helligkeit (%) | **50** | `50` |
| Gerätefunktion | **Temperatursender** | `2` |
| Senden der Helligkeit alle | **900** | `900` |
| Ausgang ist vom Typ | **1 Byte 0..100 %** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Nachlaufzeit | **60** | `60` |
| Nachlaufzeit | **60** | `60` |
| Nachlaufzeit Reduzierte Helligkeit | **30** | `30` |
| Nachlaufzeit Reduzierte Helligkeit | **30** | `30` |
| Zweistufiges Ausschalten nutzen | **ja** | `1` |
| Objekt Aktorstatus nutzen | **ja** | `1` |
| Objekt Aktorstatus nutzen | **ja** | `1` |
| Zweistufiges Ausschalten nutzen | **ja** | `1` |
| Helligkeitsunabhängige Erfassung nach Busspannungswiederkehr | **ja** | `1` |
| Objekt für helligkeitsunabhängige Erfassung nutzen | **ja** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| Eingang Slave | **ja** | `1` |
| condition parameter | **0** | `0` |
| Zusätzliche Funktionen/Objekte | **ja** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| Verwendete Helligkeit | **helligkeitsunabhängig** | `0` |
| Verwendete Helligkeit | **helligkeitsunabhängig** | `0` |
| condition parameter | **1** | `1` |
| Helligkeitsunabhängige Erfassung nach Busspannungswiederkehr | **ja** | `1` |
| condition parameter | **0** | `0` |
| Objekt für helligkeitsunabhängige Erfassung nutzen | **ja** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| Wert für reduzierte Helligkeit (%) | **50** | `50` |
| Wert für reduzierte Helligkeit (%) | **127** | `127` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| Applikation | **Aktiv** | `1` |
| Channel function table | **53253** | `53253` |
| Channel function table | **45110** | `45110` |
| condition parameter | **0** | `0` |
| Applikation | **Logik-Gatter** | `1` |
| Anzahl der Eingangsobjekte | **4** | `4` |
| logische Funktion | **OR** | `1` |
| Kanalname | **Garage-LED-Band-Gruppenstatus** | `Garage-LED-Band-Gruppenstatus` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Totzeit | **5250** | `5250` |
| Totzeit | **5250** | `5250` |
| Totzeit | **5250** | `5250` |
| condition parameter | **0** | `0` |
| Erweiterte Parameter einblenden | **ja** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| Betriebsart | **Überwachung** | `5` |
| Betriebsart | **Überwachung** | `5` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| condition parameter | **0** | `0` |
| Freigabeobjekt Melder nutzen | **ja** | `1` |
| Freigabeobjekt Melder nutzen | **ja** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| condition parameter | **1** | `1` |
| Totzeit | **5250** | `5250` |
| Filterfunktion | **aus ausfiltern** | `2` |
| Totzeit | **5250** | `5250` |
| Totzeit | **5250** | `5250` |
| logische Funktion | **OR** | `1` |
| Logik Eingang 1 | **invers** | `1` |
| Objekttyp Eingang 1 | **1 Byte** | `1` |
| Objekttyp Ausgang | **1 Byte** | `1` |
| Ausgangsobjekt senden | **bei jedem Eingangstelegramm** | `0` |
| Empfindlichkeit des Wächters | **Maximum** | `0` |
| Anzahl der Eingangsobjekte | **10** | `10` |

### ⚙️ 1.1.101 - SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten (SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten)
| Parameter | Wert | Raw |
|---|---|---|
|  | **1000** | `1000` |
|  | **10** | `10` |
| Temperature measurement | **active** | `1` |
| Send brightness on change of | **20%** | `4` |
| Detection channel (Alarm) / Direction of movement | **active** | `1` |
| Forced guidance or lock object | **force object (2Bit)** | `0` |
|     Toggle Day/Night | **directly at toggle** | `1` |
| Direction of movement | **not active** | `0` |
| Follow-up time "Day" | **1 min** | `60` |
| Output object 1 sends | **only ON** | `1` |
|     Trigger sensitivity "Day" | **6** | `6` |
| Brightness | **independent of brightness** | `1` |
| If output “Day” is already ON | **stay in automatic mode** | `0` |
|     Value for Day/Night | **Day = 0 / Night = 1** | `0` |
|     Dimming value for "Night" when ON | **100%** | `255` |
| Fallback for external push button long (Manual => Auto) | **not active** | `0` |
|     Maximum duration for short time presence | **30 s** | `30` |
| Behaviour after expiry of follow-up time | **other dimming value and switch-off delay** | `1` |
| Follow-up time "Day" | **30 s** | `30` |
|     Trigger sensitivity "Night" | **6** | `6` |
|  | **1 min** | `60` |
| Follow-up time "Night" | **1 min** | `60` |
|  | **1 min** | `60` |
| Function Buttons | **Two-button function** | `1` |
| Button assignment (left/right) | **brighter/darker** | `0` |
|     Object number | **1** | `1` |
|     Object number | **2** | `2` |
|     Send measured value on change of | **1 K** | `10` |
|     Send measured value cyclically | **30 min** | `30` |
| Request inputs after bus power return | **active** | `1` |
| Display "manual mode ON" with green LED | **not active** | `0` |
| Display "manual mode OFF" with red LED | **not active** | `0` |
| Display "lock/forced guidance ON" with green LED | **not active** | `0` |
| Display "lock/forced guidance OFF" with red LED | **not active** | `0` |
|     Trigger sensitivity "Night" | **5 (medium)** | `5` |
|     Trigger sensitivity "Day" | **5 (medium)** | `5` |
|     Switch-on threshold "Day" | **1000** | `1000` |
|  | **active** | `1` |
|  | **not active** | `0` |
|  | **not active** | `0` |
| Output object 1 sends ON cyclically | **30 s** | `30` |
| Fallback of forced guidance/lock (General setting) | **not active** | `1` |
|     Datapoint type | **1Byte DPT 5.001 Percent value (0...100%)** | `1` |
|     Percent value for output value "True" | **100%** | `255` |
| Send measured value cyclically | **15 min** | `900` |
| Active sensors | **-2** | `2` |

### ⚙️ 1.1.102 - theLuxa P300 KNX - Garage Präsenz Aussen Tor (theLuxa P300 KNX)
| Parameter | Wert | Raw |
|---|---|---|
|  Send brightness value cyclically | **every 60 min** | `10` |
|  Send temperature cyclically | **every 60 min** | `10` |
| Activate dismounting protection | **yes** | `1` |
| Use brightness sensor | **no** | `0` |
| Activate motion channel C2 | **yes..** | `1` |
| Time delay | **5 s** | `5` |
| Use temperature sensor | **yes** | `1` |
| Use brightness sensor | **no** | `0` |
| Use temperature sensor | **yes** | `1` |
| Brightness threshold value | **independent of brightness** | `0` |
| Send cyclically | **do not send cyclically** | `0` |
| Execute permanent ON | **always** | `1` |
| Program active at | **daily** | `127` |
| Activate motion channel C3 | **yes..** | `1` |
| Activate motion channel C4 | **yes..** | `1` |
| Used sensors | **centre** | `2` |
| Used sensors | **left** | `4` |
| Execute permanent ON | **always** | `1` |
| Brightness threshold value | **independent of brightness** | `0` |
| Time delay | **5 s** | `5` |
| Used sensors | **right** | `1` |
| Execute permanent ON | **always** | `1` |
| Brightness threshold value | **independent of brightness** | `0` |
| Time delay | **5 s** | `5` |
| Brightness threshold adjustable via bus | **yes** | `1` |
| Time delay adjustable via bus | **yes** | `1` |
| Time delay adjustable via bus | **yes** | `1` |
| Brightness threshold adjustable via bus | **yes** | `1` |
| Brightness threshold adjustable via bus | **yes** | `1` |

### ⚙️ 1.1.201 - Raumtemperatursensor 1-fach UP (Raumtemperatursensor 1-fach UP)
| Parameter | Wert | Raw |
|---|---|---|
| Startup delaytime | **30 s** | `30` |
| Send actual value after change of | **1,0 K** | `10` |
| Send  actual temperature cyclically | **15 min** | `15` |
| Send min/max value | **Send enable** | `1` |
| Messages | **Active** | `1` |
| Message if value > | **22 °C** | `22` |

### ⚙️ 1.3.0 - KNX IO 534 CV (4D) (KNX IO 534 CV (4D))
| Parameter | Wert | Raw |
|---|---|---|
| Device configuration | **1 x Tunable white** | `2` |
|     Channel 1 configuration | **Channel A: Cold white** | `4` |
|     Channel 2 configuration | **Channel A: Warm white** | `5` |
|     Channel 3 configuration | **Disabled** | `255` |
| RGB_COLOR_RED_USED_A_1 | **Disabled** | `0` |
| RGB_COLOR_GREEN_USED_A_1 | **Disabled** | `0` |
| RGB_COLOR_BLUE_USED_A_1 | **Disabled** | `0` |
| OUTPUT_TW_COLD_A_1 | **Enabled** | `1` |
| OUTPUT_TW_WARM_A_1 | **Enabled** | `1` |
| OUTPUT_TW_COLD_A_1 | **Enabled** | `1` |
| OUTPUT_TW_WARM_A_1 | **Enabled** | `1` |
|     Channel 3 configuration | **Disabled** | `255` |
|     Channel 4 configuration | **Disabled** | `255` |
|     Channel 4 configuration | **Disabled** | `255` |
| Send state | **On change** | `2` |
|     State objects for on/off/color temperature/
    brightness | **Enabled** | `1` |
|     State objects for cold/warm white | **Enabled** | `1` |
| Scene function | **Enabled** | `1` |
| Lock function | **Enabled** | `1` |

### ⚙️ 1.3.0 - AKD-0230CC.02 LED Controller CC/CV 30 W / 230 V 2-Kanal (AKD-0230CC.02 LED Controller CC/CV 30 W / 230 V 2-Kanal)
| Parameter | Wert | Raw |
|---|---|---|
| Device selection | **AKD-0230CC.02 LED Controller CC/CV 30 W** | `1` |
|     Town | **Berlin** | `1` |


## Detaillierte Geräteverbindungen
> **Legende**:
> * **Flags**: `K`=Kommunikation, `L`=Lesen, `S`=Schreiben, `Ü`=Übertragen, `A`=Aktualisieren
> * **Verknüpfungen**: `[S]` = Sendende Adresse (Erste Verknüpfung wenn Ü-Flag aktiv)

### 🔌 0.0.10 - Home Assistant (Dummy)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | 1 Bit (1) | `K....` | Low |  | `7/4/120` Garage-Präsenz-Innen-Externe-Bewegung<br>`0/0/4` Sommer- / Winterzeit<br>`0/0/6` Astrofunktion: Tag = 0 / Nacht = 1<br>`7/4/124` Garage-Präsenz-Innen-Meldung<br>`0/0/8` Werktag (Mo - Fr)<br>`0/1/0` Gartentor-Auf<br>`0/1/1` Gartentor-Zu<br>`0/1/2` Gartentor-Stop<br>`0/1/3` Gartentor-Status-Bewegung<br>`0/1/10` Garagentor-Fahren<br>`0/1/11` Garagentor-Stop<br>`0/1/12` Garagentor-Lüftung-Schmal<br>`0/1/17` Garagentor-Status-Offen<br>`0/1/19` Garagentor-Status-Bewegung-Öffnen<br>`0/1/20` Garagentor-Status-Bewegung-Schliessen<br>`1/0/120` Garage-Innen-LED-Band-Nachbar-Schalten<br>`1/0/121` Garage-Innen-LED-Band-Tuer-Schalten<br>`1/0/122` Garage-Innen-LED-Band-Feld-Schalten<br>`1/0/123` Garage-Innen-LED-Band-Mitte-Schalten<br>`1/0/124` Gruppe-Garage-Innen-LED-Band-Schalten<br>`1/0/125` Garage-Aussen-Downlights-Tor-Schalten<br>`1/0/126` Garage-Aussen-Downlights-Tuer-Schalten<br>`1/0/127` Gruppe-Garage-Aussen-Downlights-Schalten<br>`1/0/128` Garage-Aussen-Steckdosen-Schalten<br>`1/0/129` Garagentor-Licht-Schalten<br>`1/1/120` Garage-Innen-LED-Band-Nachbar-Status-Schalten<br>`1/1/121` Garage-Innen-LED-Band-Tuer-Status-Schalten<br>`1/1/122` Garage-Innen-LED-Band-Feld-Status-Schalten<br>`1/1/123` Garage-Innen-LED-Band-Mitte-Status-Schalten<br>`1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten<br>`1/1/125` Garage-Aussen-Downlights-Tor-Status-Schalten<br>`1/1/126` Garage-Aussen-Downlights-Tuer-Status-Schalten<br>`1/1/127` Gruppe-Garage-Aussen-Downlights-Status-Schalten<br>`1/1/128` Garage-Aussen-Steckdosen-Status-Schalten<br>`1/1/129` Garagentor-Licht-Status-Schalten<br>`0/4/1` Licht-Zentral<br>`0/4/10` Licht-Zentral-Innen<br>`0/4/12` Licht-Zentral-Innen-Unten<br>`0/4/14` Licht-Zentral-Innen-Oben<br>`0/4/20` Licht-Zentral-Aussen<br>`1/7/120` Garage-LED-Band-Dimmer-Überstrom-Alarm<br>`1/7/121` Garage-LED-Band-Dimmer-Übertemperatur-Alarm<br>`1/7/122` Garage-LED-Band-Dimmer-Spannungsversorgung<br>`0/1/50` Garage-Fenster-Auf/Ab<br>`0/1/51` Garage-Fenster-Stopp<br>`0/1/52` Garage-Fenster-Status-Bewegung<br>`7/0/14` Haus-Netzteil-Analysereset<br>`7/0/62` Garage-Netzteil-Analysereset<br>`7/0/66` Garage-Netzteil-RemoteBusreset<br>`7/0/67` Garage-Netzteil-Spannungsreset-AuxA<br>`7/4/121` Garage-Präsenz-Innen-Externer-Taster<br>`7/4/125` Garage-Präsenz-Innen-Meldung-Extern<br>`0/1/18` Garagentor-Status-Geschlossen<br>`0/1/55` Garage-Fenster-Status-aktuelle-Richtung<br>`0/1/56` Garage-Fenster-Status-Obere-Position<br>`0/1/57` Garage-Fenster-Status-Untere-Position<br>`0/1/15` Garagentor-Status-Lüftung<br>`0/1/13` Garagentor-Lüftung<br>`0/1/21` Garagentor-Status-Vorwarnzeit-Aktiv<br>`0/1/22` Garagentor-Status-Aufhaltezeit-Aktiv<br>`0/1/23` Garagentor-Status-Inspektion-Notwendig<br>`0/1/24` Garagentor-Status-Fehler-Hohe-Prio<br>`0/1/25` Garagentor-Status-Fehler-Niedrige-Prio<br>`0/1/26` Garagentor-Status-Antrieb-Gesperrt<br>`0/1/38` Garagentor-Status-Riegelstatus<br>`0/1/39` Garagentor-Status-Konfigurationsfehler<br>`0/1/40` Garagentor-Status-Torposition-Ungültig<br>`0/1/45` Alles Auf - Eingang<br>`0/1/46` Alles Auf - Ausgang 1 - Garagentor<br>`0/1/47` Alles Auf - Ausgang 2 - Gartentor<br>`0/1/14` Garagentor-Status-Lüftung-Schmal<br>`7/2/1` Regenalarm<br>`7/4/126` Garage-Präsenz-Aussen-Tor-Links<br>`7/4/127` Garage-Präsenz-Aussen-Tor-Mitte<br>`7/4/128` Garage-Präsenz-Aussen-Tor-Rechts<br>`0/0/100` NTP Server Synch anfordern |
| 1 | 1 Bit (2) | `K....` | Low |  | `1/0/10` EG-Gaestebad-Spiegel-Schalten<br>`1/0/1` EG-Flur-Deckenleuchte-Schalten<br>`1/0/2` EG-Flur-Spots-Schalten<br>`1/0/3` Flur-Treppenlicht-Schalten<br>`1/0/4` Gruppe-Flur-Schalten<br>`1/0/5` /// --- EG-Flur-Reserve --- ///<br>`1/0/6` /// --- EG-Flur-Reserve --- ///<br>`1/0/7` /// --- EG-Flur-Reserve --- ///<br>`1/0/8` /// --- EG-Flur-Reserve --- ///<br>`1/0/9` /// --- EG-Flur-Reserve --- ///<br>`1/0/11` EG-Gaestebad-Fenster-Schalten<br>`1/0/12` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/13` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/14` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/15` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/16` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/17` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/18` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/19` /// --- EG-Gaestebad-Reserve --- ///<br>`1/0/20` EG-HWR-Deckenleuchte-Schalten<br>`1/0/21` /// --- EG-HWR-Reserve --- ///<br>`1/0/22` /// --- EG-HWR-Reserve --- ///<br>`1/0/23` /// --- EG-HWR-Reserve --- ///<br>`1/0/24` /// --- EG-HWR-Reserve --- ///<br>`1/0/25` /// --- EG-HWR-Reserve --- ///<br>`1/0/26` /// --- EG-HWR-Reserve --- ///<br>`1/0/27` /// --- EG-HWR-Reserve --- ///<br>`1/0/28` /// --- EG-HWR-Reserve --- ///<br>`1/0/29` /// --- EG-HWR-Reserve --- ///<br>`1/0/30` EG-WZ-Couch-Deckenleuchte-Schalten<br>`1/0/31` EG-WZ-Wandleuchten-Schalten<br>`1/0/32` EG-WZ-Steckdosen-Schalten<br>`1/0/33` Gruppe-WZ-EZ-Steckdosen-Schalten<br>`1/0/34` /// --- EG-WZ-Reserve --- ///<br>`1/0/35` /// --- EG-WZ-Reserve --- ///<br>`1/0/36` /// --- EG-WZ-Reserve --- ///<br>`1/0/37` /// --- EG-WZ-Reserve --- ///<br>`1/0/38` /// --- EG-WZ-Reserve --- ///<br>`1/0/39` /// --- EG-WZ-Reserve --- ///<br>`1/0/40` EG-Esszimmer-Deckenleuchte-Schalten<br>`1/0/41` EG-Esszimmer-Steckdosen-Schalten<br>`1/0/42` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/43` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/44` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/45` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/46` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/47` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/48` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/49` /// --- EG-Esszimmer-Reserve --- ///<br>`1/0/50` EG-Kueche-Fenster-Spots-Schalten<br>`1/0/51` EG-Kueche-Vorn-Deckenleuchten-Schalten<br>`1/0/52` EG-Kueche-Hinten-Deckenleuchten-Schalten<br>`1/0/53` Gruppe-Kueche-Schalten<br>`1/0/54` /// --- EG-Kueche-Reserve --- ///<br>`1/0/55` /// --- EG-Kueche-Reserve --- ///<br>`1/0/56` /// --- EG-Kueche-Reserve --- ///<br>`1/0/57` /// --- EG-Kueche-Reserve --- ///<br>`1/0/58` /// --- EG-Kueche-Reserve --- ///<br>`1/0/59` /// --- EG-Kueche-Reserve --- ///<br>`1/0/60` OG-Flur-Deckenleuchte-Schalten<br>`1/0/61` OG-Flur-Wandleuchten-Schalten<br>`1/0/62` /// --- OG-Flur-Reserve --- ///<br>`1/0/63` /// --- OG-Flur-Reserve --- ///<br>`1/0/64` /// --- OG-Flur-Reserve --- ///<br>`1/0/65` /// --- OG-Flur-Reserve --- ///<br>`1/0/66` /// --- OG-Flur-Reserve --- ///<br>`1/0/67` /// --- OG-Flur-Reserve --- ///<br>`1/0/68` /// --- OG-Flur-Reserve --- ///<br>`1/0/69` /// --- OG-Flur-Reserve --- ///<br>`1/0/70` OG-Gaestezimmer-Deckenleuchte-Schalten<br>`1/0/71` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/72` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/73` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/74` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/75` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/76` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/77` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/78` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/79` /// --- OG-Gaestezimmer-Reserve --- ///<br>`1/0/80` OG-Schlafzimmer-Deckenleuchte-Schalten<br>`1/0/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Schalten<br>`1/0/82` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/83` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/84` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/85` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/86` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/87` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/88` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/89` /// --- OG-Schlafzimmer-Reserve --- ///<br>`1/0/90` OG-Arbeitszimmer-Deckenleuchte-Schalten<br>`1/0/91` OG-Arbeitszimmer-Schreibtisch-Schalten<br>`1/0/92` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/93` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/94` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/95` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/96` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/97` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/98` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/99` /// --- OG-Arbeitszimmer-Reserve --- ///<br>`1/0/100` OG-Badezimmer-Spiegel-Schalten<br>`1/0/101` OG-Badezimmer-Deckenleuchte-Schalten<br>`1/0/102` OG-Badezimmer-Spots-Schalten<br>`1/0/103` Gruppe-OG-Badezimmer-Schalten<br>`1/0/104` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/105` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/106` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/107` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/108` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/109` /// --- OG-Badezimmer-Reserve --- ///<br>`1/0/110` Haus-Aussenbeleuchtung-Schalten<br>`1/0/111` /// --- Garten-Reserve --- ///<br>`1/0/112` Haustuer-Aussenbeleuchtung-Schalten<br>`1/0/113` Garten-Weg-Vorn-Licht-Schalten<br>`1/0/114` Garten-Beet-Hinten-Licht-Schalten<br>`1/0/115` Gruppe-Garten-Wegleuchten-Schalten<br>`1/0/116` /// --- Garten-Reserve --- ///<br>`1/0/117` /// --- Garten-Reserve --- ///<br>`1/0/118` /// --- Garten-Reserve --- ///<br>`1/0/119` /// --- Garten-Reserve --- ///<br>`1/0/120` Garage-Innen-LED-Band-Nachbar-Schalten<br>`1/0/121` Garage-Innen-LED-Band-Tuer-Schalten<br>`1/0/122` Garage-Innen-LED-Band-Feld-Schalten<br>`1/0/123` Garage-Innen-LED-Band-Mitte-Schalten<br>`1/0/124` Gruppe-Garage-Innen-LED-Band-Schalten<br>`1/0/125` Garage-Aussen-Downlights-Tor-Schalten<br>`1/0/126` Garage-Aussen-Downlights-Tuer-Schalten<br>`1/0/127` Gruppe-Garage-Aussen-Downlights-Schalten<br>`1/0/128` Garage-Aussen-Steckdosen-Schalten<br>`1/0/129` Garagentor-Licht-Schalten<br>`1/0/130` /// --- Garage-Reserve --- ///<br>`1/0/131` /// --- Garage-Reserve --- ///<br>`1/0/132` /// --- Garage-Reserve --- ///<br>`1/0/133` /// --- Garage-Reserve --- ///<br>`1/0/134` /// --- Garage-Reserve --- ///<br>`1/0/135` /// --- Garage-Reserve --- ///<br>`1/0/136` /// --- Garage-Reserve --- ///<br>`1/0/137` /// --- Garage-Reserve --- ///<br>`1/0/138` /// --- Garage-Reserve --- ///<br>`1/0/139` /// --- Garage-Reserve --- ///<br>`1/0/140` Gartenhaus-Innen-Links-Schalten<br>`1/0/141` Gartenhaus-Innen-Rechts-Schalten<br>`1/0/142` Gartenhaus-Aussen-Schalten<br>`1/0/143` /// --- Gartenhaus-Reserve --- ///<br>`1/0/144` /// --- Gartenhaus-Reserve --- ///<br>`1/0/145` /// --- Gartenhaus-Reserve --- ///<br>`1/0/146` /// --- Gartenhaus-Reserve --- ///<br>`1/0/147` /// --- Gartenhaus-Reserve --- ///<br>`1/0/148` /// --- Gartenhaus-Reserve --- ///<br>`1/0/149` /// --- Gartenhaus-Reserve --- ///<br>`1/0/150` /// --- Reserve --- ///<br>`1/0/151` /// --- Reserve --- ///<br>`1/0/152` /// --- Reserve --- ///<br>`1/0/153` /// --- Reserve --- ///<br>`1/0/154` /// --- Reserve --- ///<br>`1/0/155` /// --- Reserve --- ///<br>`1/0/156` /// --- Reserve --- ///<br>`1/0/157` /// --- Reserve --- ///<br>`1/0/158` /// --- Reserve --- ///<br>`1/0/159` /// --- Reserve --- /// |
| 2 | 1 Bit (3) | `K....` | Low |  | `1/1/1` EG-Flur-Deckenleuchte-Status-Schalten<br>`1/1/2` EG-Flur-Spots-Status-Schalten<br>`1/1/3` Flur-Treppenlicht-Status-Schalten<br>`1/1/4` Gruppe-Flur-Status-Schalten<br>`1/1/10` EG-Gaestebad-Spiegel-Status-Schalten<br>`1/1/11` EG-Gaestebad-Fenster-Status-Schalten<br>`1/1/20` EG-HWR-Deckenleuchte-Status-Schalten<br>`1/1/30` EG-WZ-Couch-Deckenleuchte-Status-Schalten<br>`1/1/31` EG-WZ-Wandleuchten-Status-Schalten<br>`1/1/32` EG-WZ-Steckdosen-Status-Schalten<br>`1/1/33` Gruppe-WZ-EZ-Steckdosen-Status-Schalten<br>`1/1/40` EG-Esszimmer-Deckenleuchte-Status-Schalten<br>`1/1/41` EG-Esszimmer-Steckdosen-Status-Schalten<br>`1/1/50` EG-Kueche-Fenster-Spots-Status-Schalten<br>`1/1/51` EG-Kueche-Vorn-Deckenleuchten-Status-Schalten<br>`1/1/52` EG-Kueche-Hinten-Deckenleuchten-Status-Schalten<br>`1/1/53` Gruppe-Kueche-Status-Schalten<br>`1/1/60` OG-Flur-Deckenleuchte-Status-Schalten<br>`1/1/61` OG-Flur-Wandleuchten-Status-Schalten<br>`1/1/70` OG-Gaestezimmer-Deckenleuchte-Status-Schalten<br>`1/1/80` OG-Schlafzimmer-Deckenleuchte-Status-Schalten<br>`1/1/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Status-Schalten<br>`1/1/90` OG-Arbeitszimmer-Deckenleuchte-Status-Schalten<br>`1/1/100` OG-Badezimmer-Spiegel-Status-Schalten<br>`1/1/101` OG-Badezimmer-Deckenleuchte-Status-Schalten<br>`1/1/102` OG-Badezimmer-Spots-Status-Schalten<br>`1/1/103` Gruppe-OG-Badezimmer-Status-Schalten<br>`1/1/110` Haus-Aussenbeleuchtung-Status-Schalten<br>`1/1/113` Garten-Weg-Vorn-Licht-Status-Schalten<br>`1/1/114` Garten-Beet-Hinten-Licht-Status-Schalten<br>`0/4/2` Licht-Zentral-Status<br>`0/4/11` Licht-Zentral-Innen-Status<br>`0/4/13` Licht-Zentral-Innen-Unten-Status<br>`0/4/15` Licht-Zentral-Innen-Oben-Status<br>`0/4/21` Licht-Zentral-Aussen-Status<br>`0/2/0` Licht-Zentral-Status<br>`0/2/1` Licht-Zentral-Status<br>`0/2/2` Licht-Zentral-Status<br>`0/2/3` Licht-Zentral-Status<br>`0/2/4` Licht-Zentral-Status<br>`0/2/5` Licht-Zentral-Status<br>`0/2/6` Licht-Zentral-Status<br>`0/2/7` Licht-Zentral-Status<br>`0/2/8` Licht-Zentral-Status<br>`0/2/9` Licht-Zentral-Status |
| 3 | 1 Bit (4) | `.....` | Low |  | - |
| 4 | 1 Bit (5) | `.....` | Low |  | - |
| 5 | 1 Bit (6) | `.....` | Low |  | - |
| 6 | 1 Bit (7) | `.....` | Low |  | - |
| 7 | 1 Bit (8) | `.....` | Low |  | - |
| 8 | 1 Bit (9) | `.....` | Low |  | - |
| 9 | 1 Bit (10) | `.....` | Low |  | - |
| 10 | 2 Bit (1) | `.....` | Low |  | - |
| 11 | 2 Bit (2) | `.....` | Low |  | - |
| 12 | 2 Bit (3) | `.....` | Low |  | - |
| 13 | 3 Bit (1) | `.....` | Low |  | - |
| 14 | 3 Bit (2) | `.....` | Low |  | - |
| 15 | 3 Bit (3) | `.....` | Low |  | - |
| 16 | 4 Bit (1) | `K....` | Low |  | `1/2/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Relativ<br>`1/2/120` Garage-Innen-LED-Band-Nachbar-Dimmen-Relativ<br>`1/2/121` Garage-Innen-LED-Band-Tuer-Dimmen-Relativ<br>`1/2/122` Garage-Innen-LED-Band-Feld-Dimmen-Relativ<br>`1/2/123` Garage-Innen-LED-Band-Mitte-Dimmen-Relativ |
| 17 | 4 Bit (2) | `.....` | Low |  | - |
| 18 | 4 Bit (3) | `.....` | Low |  | - |
| 19 | 5 Bit (1) | `.....` | Low |  | - |
| 20 | 5 Bit (2) | `.....` | Low |  | - |
| 21 | 5 Bit (3) | `.....` | Low |  | - |
| 22 | 6 Bit (1) | `.....` | Low |  | - |
| 23 | 6 Bit (2) | `.....` | Low |  | - |
| 24 | 6 Bit (3) | `.....` | Low |  | - |
| 25 | 7 Bit (1) | `.....` | Low |  | - |
| 26 | 7 Bit (2) | `.....` | Low |  | - |
| 27 | 7 Bit (3) | `.....` | Low |  | - |
| 28 | 1 Byte (1) | `K....` | Low |  | `1/4/124` Gruppe-Garage-Innen-LED-Band-Status-Dimmwerrt<br>`0/1/16` Garagentor-Status-Position<br>`1/3/120` Garage-Innen-LED-Band-Nachbar-Dimmen-Absolut<br>`1/3/121` Garage-Innen-LED-Band-Tuer-Dimmen-Absolut<br>`1/3/122` Garage-Innen-LED-Band-Feld-Dimmen-Absolut<br>`1/3/123` Garage-Innen-LED-Band-Mitte-Dimmen-Absolut<br>`1/3/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Absolut<br>`1/4/120` Garage-Innen-LED-Band-Nachbar-Status-Dimmwerrt<br>`1/4/121` Garage-Innen-LED-Band-Tuer-Status-Dimmwerrt<br>`1/4/122` Garage-Innen-LED-Band-Feld-Status-Dimmwerrt<br>`1/4/123` Garage-Innen-LED-Band-Mitte-Status-Dimmwerrt<br>`0/1/53` Garage-Fenster-Status-Absolute-Position<br>`0/1/41` Garagentor-Status-Position-Invertiert |
| 29 | 1 Byte (2) | `.....` | Low |  | - |
| 30 | 1 Byte (3) | `.....` | Low |  | - |
| 31 | 1 Byte (4) | `.....` | Low |  | - |
| 32 | 1 Byte (5) | `.....` | Low |  | - |
| 33 | 1 Byte (6) | `.....` | Low |  | - |
| 34 | 1 Byte (7) | `.....` | Low |  | - |
| 35 | 1 Byte (8) | `.....` | Low |  | - |
| 36 | 1 Byte (9) | `.....` | Low |  | - |
| 37 | 1 Byte (10) | `.....` | Low |  | - |
| 38 | 2 Byte (1) | `K....` | Low |  | `7/1/121` Garage Temperatur Tür BWM<br>`7/3/121` Garage Helligkeit Tür BWM<br>`7/1/50` Gartenhaus Temperatur <br>`7/2/50` Gartenhaus Feuchtigkeit<br>`7/0/0` Haus-Netzteil-Anzahl-Spannungsausfälle<br>`7/0/3` Haus-Netzteil-Anzahl-Busresets-Bus-A<br>`7/0/6` Haus-Netzteil-Anzahl-Busresets-Bus-B<br>`7/0/9` Haus-Netzteil-AnzahlBusneustarts-Bus-A<br>`7/0/16` Haus-Netzteil-Betreibsstunden-Lebenszeit<br>`7/0/32` Haus-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde-Bus-A<br>`7/0/33` Haus-Netzteil-Mittlere-Telegrammrate-pro-Sekunde-Bus-B<br>`7/0/31` Haus-Netzteil-Temperatur<br>`7/0/23` Haus-Netzteil-Strom-Bus-A<br>`7/0/24` Haus-Netzteil-Strom-Bus-B<br>`7/0/25` Haus-Netzteil-Strom-Aux<br>`7/0/26` Haus-Netzteil-Strom-Gesamt<br>`7/0/50` Garage-Netzteil-Anzahl-Spannungsausfälle<br>`7/0/53` Garage-Netzteil-Anzahl-Busresets<br>`7/0/56` Garage-Netzteil-Anzahlspannungsreset-AuxA<br>`7/0/59` Garage-Netzteil-Anzahl-Busneustarts<br>`7/0/64` Garage-Netzteil-Betriebsstunden-Lebenszeit<br>`7/0/71` Garage-Netzteil-Strom-Bus<br>`7/0/72` Garage-Netzteil-Strom-AuxA<br>`7/0/73` Garage-Netzteil-Strom-AuxB<br>`7/0/74` Garage-Netzteil-Strom-Gesamt<br>`7/0/79` Garage-Netzteil-Temperatur<br>`7/0/80` Garage-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde<br>`7/0/81` Garage-Netzteil-Mittlere-Telegrammrate-pro-Sekunde<br>`7/3/120` Garage Helligkeit Decken BWM<br>`7/1/120` Garage Temperatur Decke BWM<br>`0/1/36` Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Stunden<br>`7/1/122` Garage Temperatur MDT Sensor<br>`7/1/1` Garage Aussen Temperatur<br>`7/3/125` Garage-Helligkeit-Aussen-Tor |
| 39 | 2 Byte (2) | `K....` | Low |  | `7/3/120` Garage Helligkeit Decken BWM<br>`7/3/121` Garage Helligkeit Tür BWM<br>`7/3/125` Garage-Helligkeit-Aussen-Tor |
| 40 | 2 Byte (3) | `.....` | Low |  | - |
| 41 | 2 Byte (4) | `.....` | Low |  | - |
| 42 | 2 Byte (5) | `.....` | Low |  | - |
| 43 | 2 Byte (6) | `.....` | Low |  | - |
| 44 | 2 Byte (7) | `.....` | Low |  | - |
| 45 | 2 Byte (8) | `.....` | Low |  | - |
| 46 | 2 Byte (9) | `.....` | Low |  | - |
| 47 | 2 Byte (10) | `.....` | Low |  | - |
| 48 | 3 Byte (1) | `K....` | Low |  | `0/0/1` Uhrzeit<br>`0/0/2` Datum<br>`7/0/1` Haus-Netzteil-LetzterSpannungsausfall-Uhrzeit<br>`7/0/2` Haus-Netzteil-LetzterSpannungsausfall-Datum<br>`7/0/4` Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-A<br>`7/0/5` Haus-Netzteil-Letzter-Busreset-Datum-Bus-A<br>`7/0/7` Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-B<br>`7/0/10` Haus-Netzteil-Letzter-Busneustart-Uhrzeit-Bus-A<br>`7/0/11` Haus-Netzteil-Letzter-Busneustart-Datum<br>`7/0/8` Haus-Netzteil-Letzter-Busreset-Datum-Bus-B<br>`7/0/51` Garage-Netzteil-LetzterSpannungsausfall-Uhrzeit<br>`7/0/52` Garage-Netzteil-LetzterSpannungsausfall-Datum<br>`7/0/54` Garage-Netzteil-Letzter-Busreset-Uhrzeit<br>`7/0/55` Garage-Netzteil-Letzter-Busreset-Datum<br>`7/0/57` Garage-Netzteil-Anzahlspannungsreset-AuxA-Uhrzeit<br>`7/0/58` Garage-Netzteil-Anzahlspannungsreset-AuxA-Datum<br>`7/0/60` Garage-Netzteil-Letzter-Busneustart-Uhrzeit<br>`7/0/61` Garage-Netzteil-Letzter-Busneustart-Datum |
| 49 | 3 Byte (2) | `.....` | Low |  | - |
| 50 | 3 Byte (3) | `.....` | Low |  | - |
| 51 | 4 Byte (1) | `K....` | Low |  | `7/0/65` Garage-Netzteil-Betriebssekunden-Lebenszeit<br>`7/0/17` Haus-Netzteil-Betriebssekunden-Lebenszeit<br>`7/0/20` Haus-Netzteil-Spannung-Bus-A<br>`7/0/21` Haus-Netzteil-Spannung-Bus-B<br>`7/0/22` Haus-Netzteil-Spannung-Aux<br>`7/0/34` Haus-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit<br>`7/0/35` Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt<br>`7/0/36` Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Reset<br>`7/0/27` Haus-Netzteil-Leistung-Bus-A<br>`7/0/28` Haus-Netzteil-Leistung-Bus-B<br>`7/0/29` Haus-Netzteil-Leistung-Aux<br>`7/0/30` Haus-Netzteil-Leistung-Gesamt<br>`7/0/37` Haus-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit<br>`7/0/38` Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt<br>`7/0/39` Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Reset<br>`7/0/68` Garage-Netzteil-Spannung-Bus<br>`7/0/69` Garage-Netzteil-Spannung-AuxA<br>`7/0/70` Garage-Netzteil-Spannung-AuxB<br>`7/0/75` Garage-Netzteil-Leistung-Bus<br>`7/0/76` Garage-Netzteil-Leistung-AuxA<br>`7/0/77` Garage-Netzteil-Leistung-AuxB<br>`7/0/78` Garage-Netzteil-Leistung-Gesamt<br>`7/0/82` Garage-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit<br>`7/0/83` Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt<br>`7/0/84` Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Analysereset<br>`7/0/85` Garage-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit<br>`7/0/86` Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt<br>`7/0/87` Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Analysereset<br>`0/1/27` Garagentor-Status-Zähler-Power-On<br>`0/1/28` Garagentor-Status-Zähler-Werks-Reset<br>`0/1/29` Garagentor-Status-Zähler-Torzyklen-Unvollst<br>`0/1/30` Garagentor-Status-Zähler-Torzyklen-Vollst<br>`0/1/31` Garagentor-Status-Zähler-Torzyklen-Unvollst-Nach-Wartung<br>`0/1/32` Garagentor-Status-Zähler-Fahrbefehle<br>`0/1/33` Garagentor-Status-Zähler-Laufzeit-Sekunden<br>`0/1/35` Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Sekunden<br>`0/1/37` Garagentor-Status-Zähler-Betriebsstunden |
| 52 | 4 Byte (2) | `.....` | Low |  | - |
| 53 | 4 Byte (3) | `.....` | Low |  | - |
| 54 | 6 Byte (1) | `.....` | Low |  | - |
| 55 | 6 Byte (2) | `.....` | Low |  | - |
| 56 | 6 Byte (3) | `.....` | Low |  | - |
| 57 | 8 Byte (1) | `K....` | Low |  | `0/0/3` Uhrzeit und Datum |
| 58 | 8 Byte (2) | `.....` | Low |  | - |
| 59 | 8 Byte (3) | `.....` | Low |  | - |
| 60 | 10 Byte (1) | `.....` | Low |  | - |
| 61 | 10 Byte (2) | `.....` | Low |  | - |
| 62 | 10 Byte (3) | `.....` | Low |  | - |
| 63 | 14 Byte (1) | `K....` | Low |  | `0/1/54` Garage-Fenster-Status-Diagnosetext |
| 64 | 14 Byte (2) | `.....` | Low |  | - |
| 65 | 14 Byte (3) | `.....` | Low |  | - |

### 🔌 1.0.9 - IP-Systemgeräte - Zusatzfunktion (IP-Systemgeräte - Zusatzfunktion)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 3 | Uhrzeit - Ausgang | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 4 | Datum - Ausgang | `K....` | Low | DPST-11-1 | `0/0/2` Datum |
| 5 | Uhrzeit und Datum - Ausgang | `K....` | Low | DPST-19-1 | `0/0/3` Uhrzeit und Datum |
| 6 | Datum/Uhrzeit - Eingang | `K....` | Low | DPST-1-17 | `0/0/5` Datum / Uhrzeit anfordern |
| 7 | NTP Server Synch. - Eingang | `K....` | Low | DPST-1-17 | `0/0/100` NTP Server Synch anfordern |
| 8 | Sommer- / Winterzeit - Ausgang | `K....` | Low | DPST-1-2 | `0/0/4` Sommer- / Winterzeit |

### 🔌 1.0.10 - Enertex KNX Dual PowerSupply 1280 (Enertex KNX Dual PowerSupply 1280)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 1 | Anzahl der Spannungsausfälle | `K....` | Low | DPST-7-1 | `7/0/0` Haus-Netzteil-Anzahl-Spannungsausfälle |
| 2 | Letzter Spannungsausfall - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/1` Haus-Netzteil-LetzterSpannungsausfall-Uhrzeit |
| 3 | Letzter Spannungsausfall - Datum | `K....` | Low | DPST-11-1 | `7/0/2` Haus-Netzteil-LetzterSpannungsausfall-Datum |
| 4 | Anzahl der Bus Resets - Bus A | `K....` | Low | DPST-7-1 | `7/0/3` Haus-Netzteil-Anzahl-Busresets-Bus-A |
| 5 | Letzter Bus Reset - Bus A - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/4` Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-A |
| 6 | Letzter Bus Reset - Bus A - Datum | `K....` | Low | DPST-11-1 | `7/0/5` Haus-Netzteil-Letzter-Busreset-Datum-Bus-A |
| 7 | Anzahl der Bus Resets - Bus B | `K....` | Low | DPST-7-1 | `7/0/6` Haus-Netzteil-Anzahl-Busresets-Bus-B |
| 8 | Letzter Bus Reset - Bus B - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/7` Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-B |
| 9 | Letzter Bus Reset - Bus B - Datum | `K....` | Low | DPST-11-1 | `7/0/8` Haus-Netzteil-Letzter-Busreset-Datum-Bus-B |
| 10 | Anzahl der Busneustarts - Bus A | `K....` | Low | DPST-7-1 | `7/0/9` Haus-Netzteil-AnzahlBusneustarts-Bus-A |
| 11 | Letzter Busneustart - Bus A - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/10` Haus-Netzteil-Letzter-Busneustart-Uhrzeit-Bus-A |
| 12 | Letzter Busneustart - Bus A - Datum | `K....` | Low | DPST-11-1 | `7/0/11` Haus-Netzteil-Letzter-Busneustart-Datum |
| 13 | Analysereset | `K....` | Low | DPST-1-17 | `7/0/14` Haus-Netzteil-Analysereset |
| 15 | Betriebsstunden Lebenszeit | `K....` | Low | DPST-7-7 | `7/0/16` Haus-Netzteil-Betreibsstunden-Lebenszeit |
| 16 | Betriebssekunden Lebenszeit | `K....` | Low | DPST-13-100 | `7/0/17` Haus-Netzteil-Betriebssekunden-Lebenszeit |
| 17 | Remote Bus Reset - Bus A | `K....` | Low | DPST-1-17 | `7/0/12` Remote-Bus-Reset-Bus-A |
| 18 | Remote Bus Reset - Bus B | `K....` | Low | DPST-1-17 | `7/0/13` Remote-Bus-Reset-Bus-B |
| 34 | Uhrzeit | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 35 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |
| 36 | Uhrzeitanfrage senden | `K....` | Low | DPST-1-17 | `0/0/5` Datum / Uhrzeit anfordern |
| 37 | Datumsanfrage senden | `K....` | Low | DPST-1-17 | `0/0/5` Datum / Uhrzeit anfordern |
| 41 | Werktag (Mo - Fr) | `K....` | Low | DPST-1-2 | `0/0/8` Werktag (Mo - Fr) |
| 50 | Spannung - Bus A | `K....` | Low | DPST-14-27 | `7/0/20` Haus-Netzteil-Spannung-Bus-A |
| 51 | Spannung - Bus B | `K....` | Low | DPST-14-27 | `7/0/21` Haus-Netzteil-Spannung-Bus-B |
| 52 | Spannung - Aux | `K....` | Low | DPST-14-27 | `7/0/22` Haus-Netzteil-Spannung-Aux |
| 53 | Strom - Bus A | `K....` | Low | DPST-14-19 | `7/0/23` Haus-Netzteil-Strom-Bus-A |
| 54 | Strom - Bus B | `K....` | Low | DPST-14-19 | `7/0/24` Haus-Netzteil-Strom-Bus-B |
| 55 | Strom - Aux | `K....` | Low | DPST-14-19 | `7/0/25` Haus-Netzteil-Strom-Aux |
| 56 | Strom - Gesamt | `K....` | Low | DPST-14-19 | `7/0/26` Haus-Netzteil-Strom-Gesamt |
| 57 | Leistung - Bus A | `K....` | Low | DPST-14-56 | `7/0/27` Haus-Netzteil-Leistung-Bus-A |
| 58 | Leistung - Bus B | `K....` | Low | DPST-14-56 | `7/0/28` Haus-Netzteil-Leistung-Bus-B |
| 59 | Leistung - Aux | `K....` | Low | DPST-14-56 | `7/0/29` Haus-Netzteil-Leistung-Aux |
| 60 | Leistung - Gesamt | `K....` | Low | DPST-14-56 | `7/0/30` Haus-Netzteil-Leistung-Gesamt |
| 61 | Temperatur | `K....` | Low | DPST-14-68 | `7/0/31` Haus-Netzteil-Temperatur |
| 62 | Aktuelle Telegrammrate (pro Sekunde) - Bus A | `K....` | Low | DPST-7-1 | `7/0/32` Haus-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde-Bus-A |
| 63 | Mittlere Telegrammrate (pro Sekunde) - Bus A | `K....` | Low | DPST-7-1 | `7/0/33` Haus-Netzteil-Mittlere-Telegrammrate-pro-Sekunde-Bus-B |
| 139 | Abgegebene Gesamtenergie Lebenszeit | `K....` | Low | DPST-13-10 | `7/0/34` Haus-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit |
| 140 | Abgegebene Gesamtenergie seit Einschaltzeitpunkt | `K....` | Low | DPST-13-10 | `7/0/35` Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt |
| 141 | Abgegebene Gesamtenergie seit letztem Analysereset | `K....` | Low | DPST-13-10 | `7/0/36` Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Reset |
| 142 | Aufgenommene Gesamtenergie Lebenszeit | `K....` | Low | DPST-13-10 | `7/0/37` Haus-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit |
| 143 | Aufgenommene Gesamtenergie seit Einschaltzeitpunkt | `K....` | Low | DPST-13-10 | `7/0/38` Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt |
| 144 | Aufgenommene Gesamtenergie seit letztem Analysereset | `K....` | Low | DPST-13-10 | `7/0/39` Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Reset |
| 150 | Tag/Nacht (1=Tag 0=Nacht) | `K....` | Low | DPST-1-2 | `0/0/6` Astrofunktion: Tag = 0 / Nacht = 1 |

### 🔌 1.0.11 - SCN-LOG1.02 Logikmodul Haus (SCN-LOG1.02 Logikmodul)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | F 1 | `K....` | Low |  | `1/1/2` EG-Flur-Spots-Status-Schalten |
| 1 | F 1 | `.....` | Low |  | - |
| 1 | F 1 | `K....` | Low |  | `1/1/3` Flur-Treppenlicht-Status-Schalten |
| 2 | F 1 | `K....` | Low |  | `1/1/61` OG-Flur-Wandleuchten-Status-Schalten |
| 9 | F 1 | `K....` | Low |  | `1/1/4` Gruppe-Flur-Status-Schalten |
| 10 | F 2 | `K....` | Low |  | `1/1/51` EG-Kueche-Vorn-Deckenleuchten-Status-Schalten |
| 11 | F 2 | `K....` | Low |  | `1/1/52` EG-Kueche-Hinten-Deckenleuchten-Status-Schalten |
| 12 | F 2 | `.....` | Low |  | - |
| 19 | F 2 | `K....` | Low |  | `1/1/53` Gruppe-Kueche-Status-Schalten |
| 20 | F 3 | `K....` | Low |  | `1/1/100` OG-Badezimmer-Spiegel-Status-Schalten |
| 21 | F 3 | `K....` | Low |  | `1/1/101` OG-Badezimmer-Deckenleuchte-Status-Schalten |
| 22 | F 3 | `K....` | Low |  | `1/1/102` OG-Badezimmer-Spots-Status-Schalten |
| 29 | F 3 | `K....` | Low |  | `1/1/103` Gruppe-OG-Badezimmer-Status-Schalten |
| 244 | Datum/Uhrzeit | `K....` | Low |  | `0/0/1` Uhrzeit |
| 245 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |

### 🔌 1.0.20 - AKU-0816.03 Universalaktor 8-fach - Haus 01 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 8 | Zentralfunktion | `KL.Ü.` | Low | DPST-1-1 | **[S]** `0/2/1` Licht-Zentral-Status<br>`0/4/1` Licht-Zentral<br>`0/4/10` Licht-Zentral-Innen |
| 31 | Kanal A | `K....` | Low | DPST-1-1 | `1/0/1` EG-Flur-Deckenleuchte-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 38 | Kanal A | `K....` | Low | DPST-1-11 | `1/1/1` EG-Flur-Deckenleuchte-Status-Schalten |
| 46 | Kanal B | `K....` | Low | DPST-1-1 | `1/0/2` EG-Flur-Spots-Schalten<br>`1/0/4` Gruppe-Flur-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 53 | Kanal B | `K....` | Low | DPST-1-11 | `1/1/2` EG-Flur-Spots-Status-Schalten |
| 61 | Kanal C | `K....` | Low | DPST-1-1 | `1/0/3` Flur-Treppenlicht-Schalten<br>`1/0/4` Gruppe-Flur-Schalten |
| 68 | Kanal C | `K....` | Low | DPST-1-11 | `1/1/3` Flur-Treppenlicht-Status-Schalten |
| 76 | Kanal D | `K....` | Low | DPST-1-1 | `1/0/60` OG-Flur-Deckenleuchte-Schalten<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 83 | Kanal D | `K....` | Low | DPST-1-11 | `1/1/60` OG-Flur-Deckenleuchte-Status-Schalten |
| 91 | Kanal E | `K....` | Low | DPST-1-1 | `1/0/61` OG-Flur-Wandleuchten-Schalten<br>`1/0/4` Gruppe-Flur-Schalten<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 98 | Kanal E | `K....` | Low | DPST-1-11 | `1/1/61` OG-Flur-Wandleuchten-Status-Schalten |
| 106 | Kanal F | `K....` | Low | DPST-1-1 | `1/0/10` EG-Gaestebad-Spiegel-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 113 | Kanal F | `K....` | Low | DPST-1-11 | `1/1/10` EG-Gaestebad-Spiegel-Status-Schalten |
| 121 | Kanal G | `K....` | Low | DPST-1-1 | `1/0/20` EG-HWR-Deckenleuchte-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 128 | Kanal G | `K....` | Low | DPST-1-11 | `1/1/20` EG-HWR-Deckenleuchte-Status-Schalten |
| 136 | Kanal H | `K....` | Low | DPST-1-1 | `1/0/70` OG-Gaestezimmer-Deckenleuchte-Schalten<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 143 | Kanal H | `K....` | Low | DPST-1-11 | `1/1/70` OG-Gaestezimmer-Deckenleuchte-Status-Schalten |

### 🔌 1.0.21 - AKU-0816.03 Universalaktor 8-fach - Haus 02 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 8 | Zentralfunktion | `KL.Ü.` | Low | DPST-1-1 | **[S]** `0/2/3` Licht-Zentral-Status<br>`0/4/1` Licht-Zentral<br>`0/4/10` Licht-Zentral-Innen |
| 31 | Kanal A | `K....` | Low | DPST-1-1 | `1/0/51` EG-Kueche-Vorn-Deckenleuchten-Schalten<br>`1/0/53` Gruppe-Kueche-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 38 | Kanal A | `K....` | Low | DPST-1-11 | `1/1/51` EG-Kueche-Vorn-Deckenleuchten-Status-Schalten |
| 46 | Kanal B | `K....` | Low | DPST-1-1 | `1/0/52` EG-Kueche-Hinten-Deckenleuchten-Schalten<br>`1/0/53` Gruppe-Kueche-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 53 | Kanal B | `K....` | Low | DPST-1-11 | `1/1/52` EG-Kueche-Hinten-Deckenleuchten-Status-Schalten |
| 61 | Kanal C | `K....` | Low | DPST-1-1 | `1/0/30` EG-WZ-Couch-Deckenleuchte-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 68 | Kanal C | `K....` | Low | DPST-1-11 | `1/1/30` EG-WZ-Couch-Deckenleuchte-Status-Schalten |
| 76 | Kanal D | `K....` | Low | DPST-1-1 | `1/0/31` EG-WZ-Wandleuchten-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 83 | Kanal D | `K....` | Low | DPST-1-11 | `1/1/31` EG-WZ-Wandleuchten-Status-Schalten |
| 91 | Kanal E | `K....` | Low | DPST-1-1 | `1/0/32` EG-WZ-Steckdosen-Schalten<br>`1/0/33` Gruppe-WZ-EZ-Steckdosen-Schalten<br>`1/0/41` EG-Esszimmer-Steckdosen-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 98 | Kanal E | `K....` | Low | DPST-1-11 | `1/1/32` EG-WZ-Steckdosen-Status-Schalten |
| 106 | Kanal F | `K....` | Low | DPST-1-1 | `1/0/40` EG-Esszimmer-Deckenleuchte-Schalten<br>`0/4/12` Licht-Zentral-Innen-Unten |
| 113 | Kanal F | `K....` | Low | DPST-1-11 | `1/1/40` EG-Esszimmer-Deckenleuchte-Status-Schalten |
| 121 | Kanal G | `K....` | Low | DPST-1-1 | `1/0/80` OG-Schlafzimmer-Deckenleuchte-Schalten<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 128 | Kanal G | `K....` | Low | DPST-1-11 | `1/1/80` OG-Schlafzimmer-Deckenleuchte-Status-Schalten |
| 136 | Kanal H | `K....` | Low | DPST-1-1 | `1/0/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Schalten<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 143 | Kanal H | `K....` | Low | DPST-1-11 | `1/1/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Status-Schalten |

### 🔌 1.0.22 - AKU-0816.03 Universalaktor 8-fach - Haus 03 (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 8 | Zentralfunktion | `KL.Ü.` | Low | DPST-1-1 | **[S]** `0/2/2` Licht-Zentral-Status<br>`0/4/1` Licht-Zentral |
| 31 | Kanal A | `K....` | Low | DPST-1-1 | `1/0/90` OG-Arbeitszimmer-Deckenleuchte-Schalten<br>`0/4/10` Licht-Zentral-Innen<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 38 | Kanal A | `K....` | Low | DPST-1-11 | `1/1/90` OG-Arbeitszimmer-Deckenleuchte-Status-Schalten |
| 46 | Kanal B | `K....` | Low | DPST-1-1 | `1/0/100` OG-Badezimmer-Spiegel-Schalten<br>`1/0/103` Gruppe-OG-Badezimmer-Schalten<br>`0/4/10` Licht-Zentral-Innen<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 53 | Kanal B | `K....` | Low | DPST-1-11 | `1/1/100` OG-Badezimmer-Spiegel-Status-Schalten |
| 61 | Kanal C | `K....` | Low | DPST-1-1 | `1/0/101` OG-Badezimmer-Deckenleuchte-Schalten<br>`1/0/103` Gruppe-OG-Badezimmer-Schalten<br>`0/4/10` Licht-Zentral-Innen<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 68 | Kanal C | `K....` | Low | DPST-1-11 | `1/1/101` OG-Badezimmer-Deckenleuchte-Status-Schalten |
| 76 | Kanal D | `K....` | Low | DPST-1-1 | `1/0/102` OG-Badezimmer-Spots-Schalten<br>`1/0/103` Gruppe-OG-Badezimmer-Schalten<br>`0/4/10` Licht-Zentral-Innen<br>`0/4/14` Licht-Zentral-Innen-Oben |
| 83 | Kanal D | `K....` | Low | DPST-1-11 | `1/1/102` OG-Badezimmer-Spots-Status-Schalten |
| 91 | Kanal E | `K....` | Low | DPST-1-1 | `1/0/110` Haus-Aussenbeleuchtung-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 98 | Kanal E | `K....` | Low | DPST-1-11 | `1/1/110` Haus-Aussenbeleuchtung-Status-Schalten |
| 106 | Kanal F | `K....` | Low | DPST-1-1 | `1/0/113` Garten-Weg-Vorn-Licht-Schalten<br>`1/0/115` Gruppe-Garten-Wegleuchten-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 113 | Kanal F | `K....` | Low | DPST-1-11 | `1/1/113` Garten-Weg-Vorn-Licht-Status-Schalten |
| 121 | Kanal G | `K....` | Low | DPST-1-1 | `1/0/114` Garten-Beet-Hinten-Licht-Schalten<br>`1/0/115` Gruppe-Garten-Wegleuchten-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 128 | Kanal G | `K....` | Low | DPST-1-11 | `1/1/114` Garten-Beet-Hinten-Licht-Status-Schalten |

### 🔌 1.0.40 - KNX Multi IO 570 (48I/O) (KNX Multi IO 570 (48I/O))
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 11 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/1` EG-Flur-Deckenleuchte-Schalten<br>`1/1/1` EG-Flur-Deckenleuchte-Status-Schalten |
| 31 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/2` EG-Flur-Spots-Schalten<br>`1/1/2` EG-Flur-Spots-Status-Schalten |
| 51 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/3` Flur-Treppenlicht-Schalten<br>`1/1/3` Flur-Treppenlicht-Status-Schalten |
| 71 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/60` OG-Flur-Deckenleuchte-Schalten<br>`1/1/60` OG-Flur-Deckenleuchte-Status-Schalten |
| 91 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/61` OG-Flur-Wandleuchten-Schalten<br>`1/1/61` OG-Flur-Wandleuchten-Status-Schalten |
| 111 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/10` EG-Gaestebad-Spiegel-Schalten<br>`1/1/10` EG-Gaestebad-Spiegel-Status-Schalten |
| 131 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/20` EG-HWR-Deckenleuchte-Schalten<br>`1/1/20` EG-HWR-Deckenleuchte-Status-Schalten |
| 151 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/52` EG-Kueche-Hinten-Deckenleuchten-Schalten<br>`1/1/52` EG-Kueche-Hinten-Deckenleuchten-Status-Schalten |
| 171 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/51` EG-Kueche-Vorn-Deckenleuchten-Schalten<br>`1/1/51` EG-Kueche-Vorn-Deckenleuchten-Status-Schalten |
| 191 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/40` EG-Esszimmer-Deckenleuchte-Schalten<br>`1/1/40` EG-Esszimmer-Deckenleuchte-Status-Schalten |
| 211 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/30` EG-WZ-Couch-Deckenleuchte-Schalten<br>`1/1/30` EG-WZ-Couch-Deckenleuchte-Status-Schalten |
| 231 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/31` EG-WZ-Wandleuchten-Schalten<br>`1/1/31` EG-WZ-Wandleuchten-Status-Schalten |
| 251 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/32` EG-WZ-Steckdosen-Schalten<br>`1/0/33` Gruppe-WZ-EZ-Steckdosen-Schalten<br>`1/1/32` EG-WZ-Steckdosen-Status-Schalten |
| 271 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/70` OG-Gaestezimmer-Deckenleuchte-Schalten<br>`1/1/70` OG-Gaestezimmer-Deckenleuchte-Status-Schalten |
| 291 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/102` OG-Badezimmer-Spots-Schalten<br>`1/1/102` OG-Badezimmer-Spots-Status-Schalten |
| 311 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/101` OG-Badezimmer-Deckenleuchte-Schalten<br>`1/1/101` OG-Badezimmer-Deckenleuchte-Status-Schalten |
| 331 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/100` OG-Badezimmer-Spiegel-Schalten<br>`1/1/100` OG-Badezimmer-Spiegel-Status-Schalten |
| 351 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/90` OG-Arbeitszimmer-Deckenleuchte-Schalten<br>`1/1/90` OG-Arbeitszimmer-Deckenleuchte-Status-Schalten |
| 371 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/80` OG-Schlafzimmer-Deckenleuchte-Schalten<br>`1/1/80` OG-Schlafzimmer-Deckenleuchte-Status-Schalten |
| 391 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Schalten<br>`1/1/81` OG-Schlafzimmer-Ankleide-Deckenleuchte-Status-Schalten |
| 411 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `0/4/1` Licht-Zentral |
| 431 | GO_BASE_1 | `KLS..` | Low | DPT-16 DPST-16-0 | `0/4/20` Licht-Zentral-Aussen |
| 451 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `0/4/1` Licht-Zentral |
| 471 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/53` Gruppe-Kueche-Schalten<br>`1/1/53` Gruppe-Kueche-Status-Schalten |
| 491 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/103` Gruppe-OG-Badezimmer-Schalten<br>`1/1/103` Gruppe-OG-Badezimmer-Status-Schalten |
| 511 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/4` Gruppe-Flur-Schalten<br>`1/1/4` Gruppe-Flur-Status-Schalten |
| 531 | GO_BASE_1 | `K.S..` | Low | DPT-16 DPST-16-0 | `1/0/91` OG-Arbeitszimmer-Schreibtisch-Schalten |
| 973 | GO_BASE_33 | `.L...` | Low | DPT-1 DPST-1-2 | - |

### 🔌 1.1.10 - Enertex KNX PowerSupply 960³ (Enertex KNX PowerSupply 960³)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 1 | Anzahl der Spannungsausfälle | `K....` | Low | DPST-7-1 | `7/0/50` Garage-Netzteil-Anzahl-Spannungsausfälle |
| 2 | Letzter Spannungsausfall - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/51` Garage-Netzteil-LetzterSpannungsausfall-Uhrzeit |
| 3 | Letzter Spannungsausfall - Datum | `K....` | Low | DPST-11-1 | `7/0/52` Garage-Netzteil-LetzterSpannungsausfall-Datum |
| 4 | Anzahl der Bus Resets - Bus A | `K....` | Low | DPST-7-1 | `7/0/53` Garage-Netzteil-Anzahl-Busresets |
| 5 | Letzter Bus Reset - Bus A - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/54` Garage-Netzteil-Letzter-Busreset-Uhrzeit |
| 6 | Letzter Bus Reset - Bus A - Datum | `K....` | Low | DPST-11-1 | `7/0/55` Garage-Netzteil-Letzter-Busreset-Datum |
| 7 | Anzahl der Bus Resets - Bus B | `K....` | Low | DPST-7-1 | `7/0/56` Garage-Netzteil-Anzahlspannungsreset-AuxA |
| 8 | Letzter Bus Reset - Bus B - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/57` Garage-Netzteil-Anzahlspannungsreset-AuxA-Uhrzeit |
| 9 | Letzter Bus Reset - Bus B - Datum | `K....` | Low | DPST-11-1 | `7/0/58` Garage-Netzteil-Anzahlspannungsreset-AuxA-Datum |
| 10 | Anzahl der Busneustarts - Bus A | `K....` | Low | DPST-7-1 | `7/0/59` Garage-Netzteil-Anzahl-Busneustarts |
| 11 | Letzter Busneustart - Bus A - Uhrzeit | `K....` | Low | DPST-10-1 | `7/0/60` Garage-Netzteil-Letzter-Busneustart-Uhrzeit |
| 12 | Letzter Busneustart - Bus A - Datum | `K....` | Low | DPST-11-1 | `7/0/61` Garage-Netzteil-Letzter-Busneustart-Datum |
| 13 | Analysereset | `K....` | Low | DPST-1-17 | `7/0/62` Garage-Netzteil-Analysereset |
| 15 | Betriebsstunden Lebenszeit | `K....` | Low | DPST-7-7 | `7/0/64` Garage-Netzteil-Betriebsstunden-Lebenszeit |
| 16 | Betriebssekunden Lebenszeit | `K....` | Low | DPST-13-100 | `7/0/65` Garage-Netzteil-Betriebssekunden-Lebenszeit |
| 17 | Remote Bus Reset - Bus A | `K....` | Low | DPST-1-17 | `7/0/66` Garage-Netzteil-RemoteBusreset |
| 18 | Remote Bus Reset - Bus B | `K....` | Low | DPST-1-17 | `7/0/67` Garage-Netzteil-Spannungsreset-AuxA |
| 34 | Uhrzeit | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 35 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |
| 36 | Uhrzeitanfrage senden | `K....` | Low | DPST-1-17 | `0/0/5` Datum / Uhrzeit anfordern |
| 37 | Datumsanfrage senden | `K....` | Low | DPST-1-17 | `0/0/5` Datum / Uhrzeit anfordern |
| 41 | Werktag (Mo - Fr) | `K....` | Low | DPST-1-2 | `0/0/8` Werktag (Mo - Fr) |
| 50 | Spannung - Bus A | `K....` | Low | DPST-14-27 | `7/0/68` Garage-Netzteil-Spannung-Bus |
| 51 | Spannung - Bus B | `K....` | Low | DPST-14-27 | `7/0/69` Garage-Netzteil-Spannung-AuxA |
| 52 | Spannung - Aux | `K....` | Low | DPST-14-27 | `7/0/70` Garage-Netzteil-Spannung-AuxB |
| 53 | Strom - Bus A | `K....` | Low | DPST-14-19 | `7/0/71` Garage-Netzteil-Strom-Bus |
| 54 | Strom - Bus B | `K....` | Low | DPST-14-19 | `7/0/72` Garage-Netzteil-Strom-AuxA |
| 55 | Strom - Aux | `K....` | Low | DPST-14-19 | `7/0/73` Garage-Netzteil-Strom-AuxB |
| 56 | Strom - Gesamt | `K....` | Low | DPST-14-19 | `7/0/74` Garage-Netzteil-Strom-Gesamt |
| 57 | Leistung - Bus A | `K....` | Low | DPST-14-56 | `7/0/75` Garage-Netzteil-Leistung-Bus |
| 58 | Leistung - Bus B | `K....` | Low | DPST-14-56 | `7/0/76` Garage-Netzteil-Leistung-AuxA |
| 59 | Leistung - Aux | `K....` | Low | DPST-14-56 | `7/0/77` Garage-Netzteil-Leistung-AuxB |
| 60 | Leistung - Gesamt | `K....` | Low | DPST-14-56 | `7/0/78` Garage-Netzteil-Leistung-Gesamt |
| 61 | Temperatur | `K....` | Low | DPST-14-68 | `7/0/79` Garage-Netzteil-Temperatur |
| 62 | Aktuelle Telegrammrate (pro Sekunde) - Bus A | `K....` | Low | DPST-7-1 | `7/0/80` Garage-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde |
| 63 | Mittlere Telegrammrate (pro Sekunde) - Bus A | `K....` | Low | DPST-7-1 | `7/0/81` Garage-Netzteil-Mittlere-Telegrammrate-pro-Sekunde |
| 139 | Abgegebene Gesamtenergie Lebenszeit | `K....` | Low | DPST-13-10 | `7/0/82` Garage-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit |
| 140 | Abgegebene Gesamtenergie seit Einschaltzeitpunkt | `K....` | Low | DPST-13-10 | `7/0/83` Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt |
| 141 | Abgegebene Gesamtenergie seit letztem Analysereset | `K....` | Low | DPST-13-10 | `7/0/84` Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Analysereset |
| 142 | Aufgenommene Gesamtenergie Lebenszeit | `K....` | Low | DPST-13-10 | `7/0/85` Garage-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit |
| 143 | Aufgenommene Gesamtenergie seit Einschaltzeitpunkt | `K....` | Low | DPST-13-10 | `7/0/86` Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt |
| 144 | Aufgenommene Gesamtenergie seit letztem Analysereset | `K....` | Low | DPST-13-10 | `7/0/87` Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Analysereset |
| 150 | Tag/Nacht (1=Tag 0=Nacht) | `K....` | Low | DPST-1-2 | `0/0/6` Astrofunktion: Tag = 0 / Nacht = 1 |
| 178 | Schaltuhr 1 - Telegramm 1 | `K....` | Low | DPST-232-600 | `0/4/1` Licht-Zentral |
| 187 | Schaltuhr 2 - Telegramm 1 | `K....` | Low | DPST-232-600 | `0/4/20` Licht-Zentral-Aussen |
| 188 | Schaltuhr 2 - Telegramm 2 | `K....` | Low | DPST-232-600 | `0/4/20` Licht-Zentral-Aussen |
| 196 | Schaltuhr 3 - Telegramm 1 | `K....` | Low | DPST-232-600 | `0/4/20` Licht-Zentral-Aussen |
| 197 | Schaltuhr 3 - Telegramm 2 | `K....` | Low | DPST-232-600 | `0/4/20` Licht-Zentral-Aussen |

### 🔌 1.1.11 - SCN-LOG1.02 Logikmodul Garage (SCN-LOG1.02 Logikmodul)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | F 1 | `K....` | Low |  | `0/1/45` Alles Auf - Eingang |
| 1 | F 1 | `K....` | Low |  | `0/1/46` Alles Auf - Ausgang 1 - Garagentor |
| 2 | F 1 | `K....` | Low |  | `0/1/47` Alles Auf - Ausgang 2 - Gartentor |
| 10 | F 2 | `K....` | Low |  | `1/4/120` Garage-Innen-LED-Band-Nachbar-Status-Dimmwerrt |
| 11 | F 2 | `K....` | Low |  | `1/4/121` Garage-Innen-LED-Band-Tuer-Status-Dimmwerrt |
| 12 | F 2 | `K....` | Low |  | `1/4/122` Garage-Innen-LED-Band-Feld-Status-Dimmwerrt |
| 13 | F 2 | `K....` | Low |  | `1/4/123` Garage-Innen-LED-Band-Mitte-Status-Dimmwerrt |
| 14 | F 2 | `K....` | Low |  | `1/4/124` Gruppe-Garage-Innen-LED-Band-Status-Dimmwerrt |
| 29 | F 3 | `.....` | Low |  | - |
| 244 | Datum/Uhrzeit | `K....` | Low |  | `0/0/1` Uhrzeit |
| 245 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |

### 🔌 1.1.20 - AKU-0816.03 Universalaktor 8-fach - Garage (AKU-0816.03 Universalaktor 8-fach, 16A, 230VAC)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 8 | Zentralfunktion | `KL.Ü.` | Low | DPST-1-1 | **[S]** `0/2/0` Licht-Zentral-Status<br>`0/4/1` Licht-Zentral |
| 31 | Kanal A / B | `K....` | Low | DPST-1-8 | `0/1/50` Garage-Fenster-Auf/Ab |
| 33 | Kanal A / B | `K....` | Low | DPST-1-17 | `0/1/51` Garage-Fenster-Stopp |
| 35 | Kanal A / B | `K....` | Low | DPST-1-8 | `0/1/55` Garage-Fenster-Status-aktuelle-Richtung |
| 36 | Kanal A / B | `K....` | Low | DPST-1-11 | `0/1/52` Garage-Fenster-Status-Bewegung |
| 40 | Kanal A / B | `K....` | Low | DPST-5-1 | `0/1/53` Garage-Fenster-Status-Absolute-Position |
| 45 | Kanal A / B | `K....` | Low | DPST-1-11 | `0/1/56` Garage-Fenster-Status-Obere-Position |
| 46 | Kanal A / B | `K....` | Low | DPST-1-11 | `0/1/57` Garage-Fenster-Status-Untere-Position |
| 51 | Kanal A / B | `K....` | Low | DPST-1-5 | `7/2/1` Regenalarm |
| 60 | Kanal A / B | `K....` | Low | DPST-16-0 | `0/1/54` Garage-Fenster-Status-Diagnosetext |
| 61 | Kanal C | `K....` | Low | DPST-1-1 | `1/0/125` Garage-Aussen-Downlights-Tor-Schalten<br>`1/0/127` Gruppe-Garage-Aussen-Downlights-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 68 | Kanal C | `K....` | Low | DPST-1-11 | `1/1/125` Garage-Aussen-Downlights-Tor-Status-Schalten |
| 76 | Kanal D | `K....` | Low | DPST-1-1 | `1/0/126` Garage-Aussen-Downlights-Tuer-Schalten<br>`1/0/127` Gruppe-Garage-Aussen-Downlights-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 83 | Kanal D | `K....` | Low | DPST-1-11 | `1/1/126` Garage-Aussen-Downlights-Tuer-Status-Schalten |
| 91 | Kanal E | `K....` | Low | DPST-1-1 | `1/0/128` Garage-Aussen-Steckdosen-Schalten<br>`0/4/20` Licht-Zentral-Aussen |
| 98 | Kanal E | `K....` | Low | DPST-1-11 | `1/1/128` Garage-Aussen-Steckdosen-Status-Schalten |

### 🔌 1.1.21 - AKD-0424R2.02 - Dimmaktor Garage (AKD-0424R2.02 LED Controller 4 Kanal/RGBW, 2TE REG)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Kanal A | `K....` | Low | DPST-1-1 | `1/0/120` Garage-Innen-LED-Band-Nachbar-Schalten |
| 2 | Kanal A | `K....` | Low | DPST-3-7 | `1/2/120` Garage-Innen-LED-Band-Nachbar-Dimmen-Relativ |
| 3 | Kanal A | `K....` | Low | DPST-5-1 | `1/3/120` Garage-Innen-LED-Band-Nachbar-Dimmen-Absolut |
| 4 | Kanal A | `K....` | Low | DPST-1-11 | `1/1/120` Garage-Innen-LED-Band-Nachbar-Status-Schalten |
| 5 | Kanal A | `K....` | Low | DPST-5-1 | `1/4/120` Garage-Innen-LED-Band-Nachbar-Status-Dimmwerrt |
| 16 | Kanal B | `K....` | Low | DPST-1-1 | `1/0/121` Garage-Innen-LED-Band-Tuer-Schalten |
| 18 | Kanal B | `K....` | Low | DPST-3-7 | `1/2/121` Garage-Innen-LED-Band-Tuer-Dimmen-Relativ |
| 19 | Kanal B | `K....` | Low | DPST-5-1 | `1/3/121` Garage-Innen-LED-Band-Tuer-Dimmen-Absolut |
| 20 | Kanal B | `K....` | Low | DPST-1-11 | `1/1/121` Garage-Innen-LED-Band-Tuer-Status-Schalten |
| 21 | Kanal B | `K....` | Low | DPST-5-1 | `1/4/121` Garage-Innen-LED-Band-Tuer-Status-Dimmwerrt |
| 32 | Kanal C | `K....` | Low | DPST-1-1 | `1/0/122` Garage-Innen-LED-Band-Feld-Schalten |
| 34 | Kanal C | `K....` | Low | DPST-3-7 | `1/2/122` Garage-Innen-LED-Band-Feld-Dimmen-Relativ |
| 35 | Kanal C | `K....` | Low | DPST-5-1 | `1/3/122` Garage-Innen-LED-Band-Feld-Dimmen-Absolut |
| 36 | Kanal C | `K....` | Low | DPST-1-11 | `1/1/122` Garage-Innen-LED-Band-Feld-Status-Schalten |
| 37 | Kanal C | `K....` | Low | DPST-5-1 | `1/4/122` Garage-Innen-LED-Band-Feld-Status-Dimmwerrt |
| 48 | Kanal D | `K....` | Low | DPST-1-1 | `1/0/123` Garage-Innen-LED-Band-Mitte-Schalten |
| 50 | Kanal D | `K....` | Low | DPST-3-7 | `1/2/123` Garage-Innen-LED-Band-Mitte-Dimmen-Relativ |
| 51 | Kanal D | `K....` | Low | DPST-5-1 | `1/3/123` Garage-Innen-LED-Band-Mitte-Dimmen-Absolut |
| 52 | Kanal D | `K....` | Low | DPST-1-11 | `1/1/123` Garage-Innen-LED-Band-Mitte-Status-Schalten |
| 53 | Kanal D | `K....` | Low | DPST-5-1 | `1/4/123` Garage-Innen-LED-Band-Mitte-Status-Dimmwerrt |
| 135 | Zentral | `KL.Ü.` | Low | DPST-1-1 | **[S]** `0/2/4` Licht-Zentral-Status<br>`1/0/124` Gruppe-Garage-Innen-LED-Band-Schalten<br>`0/4/1` Licht-Zentral<br>`0/4/10` Licht-Zentral-Innen |
| 136 | Zentral | `K....` | Low | DPST-3-7 | `1/2/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Relativ |
| 137 | Zentral | `K....` | Low | DPST-5-1 | `1/3/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Absolut |
| 139 | Zentral | `K....` | High | DPST-1-5 | `1/7/120` Garage-LED-Band-Dimmer-Überstrom-Alarm |
| 140 | Zentral | `K....` | High | DPST-1-5 | `1/7/121` Garage-LED-Band-Dimmer-Übertemperatur-Alarm |
| 143 | Zentral | `K....` | Low | DPST-1-11 | `1/7/122` Garage-LED-Band-Dimmer-Spannungsversorgung |
| 144 | Uhrzeit | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 145 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |
| 146 | Datum/Uhrzeit | `K....` | Low | DPST-19-1 | `0/0/3` Uhrzeit und Datum |
| 148 | Tag / Nacht | `K....` | Low | DPST-1-2 | `0/0/6` Astrofunktion: Tag = 0 / Nacht = 1 |

### 🔌 1.1.30 - Hörmann KNX-Gateway (KNX-Gateway)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 1 | - | `K....` | Low |  | `0/1/10` Garagentor-Fahren<br>`0/1/46` Alles Auf - Ausgang 1 - Garagentor |
| 2 | - | `K....` | Low |  | `0/1/11` Garagentor-Stop |
| 3 | - | `K....` | Low |  | `0/1/12` Garagentor-Lüftung-Schmal |
| 5 | - | `K....` | Low |  | `0/1/13` Garagentor-Lüftung |
| 7 | - | `K....` | Low |  | `1/0/129` Garagentor-Licht-Schalten |
| 11 | - | `K....` | Low |  | `0/1/17` Garagentor-Status-Offen |
| 12 | - | `K....` | Low |  | `0/1/18` Garagentor-Status-Geschlossen |
| 13 | - | `K....` | Low |  | `0/1/14` Garagentor-Status-Lüftung-Schmal |
| 14 | - | `K....` | Low |  | `0/1/16` Garagentor-Status-Position |
| 15 | - | `K....` | Low |  | `0/1/19` Garagentor-Status-Bewegung-Öffnen |
| 16 | - | `K....` | Low |  | `0/1/20` Garagentor-Status-Bewegung-Schliessen |
| 17 | - | `K....` | Low |  | `1/1/129` Garagentor-Licht-Status-Schalten |
| 24 | - | `K....` | Low |  | `0/1/21` Garagentor-Status-Vorwarnzeit-Aktiv |
| 25 | - | `K....` | Low |  | `0/1/22` Garagentor-Status-Aufhaltezeit-Aktiv |
| 26 | - | `K....` | Low |  | `0/1/23` Garagentor-Status-Inspektion-Notwendig |
| 27 | - | `K....` | Low |  | `0/1/24` Garagentor-Status-Fehler-Hohe-Prio |
| 28 | - | `K....` | Low |  | `0/1/25` Garagentor-Status-Fehler-Niedrige-Prio |
| 29 | - | `K....` | Low |  | `0/1/26` Garagentor-Status-Antrieb-Gesperrt |
| 30 | - | `K....` | Low |  | `0/1/15` Garagentor-Status-Lüftung |
| 31 | - | `K....` | Low |  | `0/1/27` Garagentor-Status-Zähler-Power-On |
| 32 | - | `K....` | Low |  | `0/1/28` Garagentor-Status-Zähler-Werks-Reset |
| 34 | - | `K....` | Low |  | `0/1/29` Garagentor-Status-Zähler-Torzyklen-Unvollst |
| 35 | - | `K....` | Low |  | `0/1/30` Garagentor-Status-Zähler-Torzyklen-Vollst |
| 36 | - | `K....` | Low |  | `0/1/31` Garagentor-Status-Zähler-Torzyklen-Unvollst-Nach-Wartung |
| 37 | - | `K....` | Low |  | `0/1/32` Garagentor-Status-Zähler-Fahrbefehle |
| 38 | - | `K....` | Low |  | `0/1/33` Garagentor-Status-Zähler-Laufzeit-Sekunden |
| 39 | - | `K....` | Low |  | `0/1/34` Garagentor-Status-Zähler-Laufzeit-Stunden |
| 40 | - | `K....` | Low |  | `0/1/35` Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Sekunden |
| 41 | - | `K....` | Low |  | `0/1/36` Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Stunden |
| 42 | - | `K....` | Low |  | `0/1/37` Garagentor-Status-Zähler-Betriebsstunden |
| 46 | - | `K....` | Low |  | `0/1/38` Garagentor-Status-Riegelstatus |
| 47 | - | `K....` | Low |  | `0/1/39` Garagentor-Status-Konfigurationsfehler |
| 48 | - | `K....` | Low |  | `0/1/40` Garagentor-Status-Torposition-Ungültig |

### 🔌 1.1.50 - BE-GT20x.02 Glastaster II Smart (BE-GT20x.02 Glastaster II Smart)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | F1: Taste 1 | `K....` | Low |  | `0/1/45` Alles Auf - Eingang |
| 1 | F1: Taste 1 | `K....` | Low |  | `0/1/17` Garagentor-Status-Offen |
| 7 | F2: Taste 1 | `K....` | Low |  | `0/1/0` Gartentor-Auf |
| 10 | F2: Taste 1 | `K....` | Low |  | `0/1/3` Gartentor-Status-Bewegung |
| 14 | F3: Taste 1 | `K....` | Low |  | `0/1/13` Garagentor-Lüftung |
| 15 | F3: Taste 1 | `K....` | Low |  | `0/1/15` Garagentor-Status-Lüftung |
| 21 | F4: Taste 1 | `K....` | Low |  | `0/1/50` Garage-Fenster-Auf/Ab |
| 22 | F4: Taste 1 | `K....` | Low |  | `0/1/51` Garage-Fenster-Stopp |
| 24 | F4: Taste 1 | `K....` | Low |  | `0/1/53` Garage-Fenster-Status-Absolute-Position |
| 28 | F5/6: Taste 1 | `K....` | Low |  | `0/1/10` Garagentor-Fahren |
| 29 | F5/6: Taste 1 | `K....` | Low |  | `0/1/11` Garagentor-Stop |
| 31 | F5/6 lang: Taste 1 | `K....` | Low |  | `0/1/16` Garagentor-Status-Position |
| 42 | F7/8: Taste 1 | `K....` | Low |  | `7/4/121` Garage-Präsenz-Innen-Externer-Taster |
| 42 | F7/8: Taste 1 | `.LS..` | Low |  | - |
| 43 | F7/8: Taste 1 | `K....` | Low |  | `1/2/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Relativ |
| 45 | F7/8 lang: Taste 1 | `K....` | Low |  | `1/4/124` Gruppe-Garage-Innen-LED-Band-Status-Dimmwerrt |
| 84 | Patsch-Taste kurz | `K....` | Low |  | `1/0/124` Gruppe-Garage-Innen-LED-Band-Schalten |
| 85 | Patsch-Taste kurz | `K....` | Low |  | `1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten |
| 100 | Logik 4 | `K....` | Low | DPST-1-1 | `1/1/125` Garage-Aussen-Downlights-Tor-Status-Schalten |
| 101 | Logik 4 | `K....` | Low | DPST-1-1 | `1/1/126` Garage-Aussen-Downlights-Tuer-Status-Schalten |
| 102 | Logik 4 | `K....` | Low |  | `1/1/127` Gruppe-Garage-Aussen-Downlights-Status-Schalten |
| 103 | LED 1 | `K....` | Low |  | `0/1/19` Garagentor-Status-Bewegung-Öffnen<br>`0/1/20` Garagentor-Status-Bewegung-Schliessen |
| 104 | LED 2 | `K....` | Low |  | `0/1/3` Gartentor-Status-Bewegung |
| 105 | LED 3 | `K....` | Low |  | `0/1/15` Garagentor-Status-Lüftung |
| 106 | LED 4 | `K....` | Low |  | `0/1/56` Garage-Fenster-Status-Obere-Position |
| 107 | LED 5 | `K....` | Low |  | `0/1/17` Garagentor-Status-Offen |
| 108 | LED 6 | `K....` | Low |  | `0/1/17` Garagentor-Status-Offen |
| 109 | LED 7 | `K....` | Low |  | `1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten |
| 110 | LED 8 | `K....` | Low |  | `1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten |
| 133 | Tag / Nacht | `K....` | Low | DPST-1-2 | `0/0/6` Astrofunktion: Tag = 0 / Nacht = 1 |
| 134 | Präsenz | `K....` | Low | DPST-1-1 | `7/4/124` Garage-Präsenz-Innen-Meldung |
| 139 | Uhrzeit | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 140 | Datum | `K....` | Low | DPST-11-1 | `0/0/2` Datum |
| 141 | Uhrzeit/Datum | `K....` | Low | DPST-19-1 | `0/0/3` Uhrzeit und Datum |
| 149 | Statuswert 1 | `K....` | Low |  | `7/1/50` Gartenhaus Temperatur  |
| 150 | Statuswert 2 | `K....` | Low |  | `7/2/50` Gartenhaus Feuchtigkeit |

### 🔌 1.1.51 - Garage Taster Tor-Nachbar (Taster Busankoppler AP WG 2fach Zweipunktbedienung)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Taste 1 | `K.S..` | Low | DPT-5 | `0/1/10` Garagentor-Fahren<br>`0/1/17` Garagentor-Status-Offen |
| 2 | Taste 3 | `K....` | Low | DPT-5 | `0/1/0` Gartentor-Auf<br>`0/1/3` Gartentor-Status-Bewegung |
| 3 | Taste 4 | `K....` | Low | DPT-5 | `0/1/0` Gartentor-Auf |

### 🔌 1.1.52 - Garage Taster Tor-Feld (Taster Busankoppler AP WG 2fach Zweipunktbedienung)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Taste 1 | `K.S..` | Low | DPT-5 | `0/1/10` Garagentor-Fahren<br>`0/1/17` Garagentor-Status-Offen |
| 2 | Taste 3 | `K....` | Low | DPT-5 | `0/1/0` Gartentor-Auf<br>`0/1/3` Gartentor-Status-Bewegung |
| 3 | Taste 4 | `K....` | Low | DPT-5 | `0/1/0` Gartentor-Auf |

### 🔌 1.1.70 - KNX IO 410 - Binäreingang 4-fach Garage (KNX IO 410 (4I))
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 11 | GO_BASE_11 | `..S..` | Low | DPT-16 DPST-16-0 | - |
| 11 | GO_BASE_11 | `KLS..` | Low | DPT-16 DPST-16-0 | `7/2/1` Regenalarm |
| 12 | GO_BASE_12 | `..S..` | Low | DPT-5 DPST-5-1 | - |
| 31 | GO_BASE_21 | `..S..` | Low | DPT-16 DPST-16-0 | - |
| 32 | GO_BASE_22 | `..S..` | Low | DPT-5 DPST-5-1 | - |
| 51 | GO_BASE_31 | `..S..` | Low | DPT-16 DPST-16-0 | - |
| 71 | GO_BASE_41 | `..S..` | Low | DPT-16 DPST-16-0 | - |

### 🔌 1.1.100 - 6131/31 Busch-Präsenzmelder Premium (6131/31 Busch-Präsenzmelder Premium)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 3 | P1: Slave | `K....` | Low | DPST-1-1 | `7/4/120` Garage-Präsenz-Innen-Externe-Bewegung |
| 4 | P1: Aktorstatus | `K....` | Low | DPST-3-7 | `1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten |
| 10 | P1: Bewegung (Master) | `K....` | Low | DPST-7-1 | `1/3/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Absolut |
| 13 | P1: Externer Taster | `K....` | Low | DPT-9 | `7/4/121` Garage-Präsenz-Innen-Externer-Taster |
| 55 | P4: Slave | `K....` | Low | DPST-1-1 | `7/4/125` Garage-Präsenz-Innen-Meldung-Extern |
| 62 | P4: Bewegung (Master) | `KL...` | Low | DPST-5-1 | `7/4/124` Garage-Präsenz-Innen-Meldung |
| 69 | BR: Helligkeit | `KL...` | Low | DPT-9 | `7/3/120` Garage Helligkeit Decken BWM |
| 82 | RTC: Ist-Temperatur | `KL...` | Low | DPT-9 | `7/1/120` Garage Temperatur Decke BWM |
| 244 | GF1: Ausgang | `K....` | Low |  | `1/1/124` Gruppe-Garage-Innen-LED-Band-Status-Schalten |
| 245 | GF1: Eingang 1 | `K....` | Low |  | `1/1/120` Garage-Innen-LED-Band-Nachbar-Status-Schalten |
| 246 | GF1: Eingang 2 | `K....` | Low | DPST-7-1 | `1/1/121` Garage-Innen-LED-Band-Tuer-Status-Schalten |
| 247 | GF1: Eingang 3 | `K....` | Low | DPST-5-1 | `1/1/122` Garage-Innen-LED-Band-Feld-Status-Schalten |
| 248 | GF1: Eingang 4 | `K....` | Low | DPST-5-1 | `1/1/123` Garage-Innen-LED-Band-Mitte-Status-Schalten |

### 🔌 1.1.101 - SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten (SCN-BWM55T.G2 Bewegungsmelder TS 55 mit T-Sensor und 2 Sensortasten)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Lichtkanal 1 - Ausgang 1 | `K....` | Low |  | `7/4/120` Garage-Präsenz-Innen-Externe-Bewegung |
| 75 | Alarm - Ausgang | `K....` | Low | DPST-1-1 | `7/4/125` Garage-Präsenz-Innen-Meldung-Extern |
| 90 | Tag/Nacht | `K....` | Low | DPST-1-2 | `0/0/6` Astrofunktion: Tag = 0 / Nacht = 1 |
| 97 | Helligkeit | `K....` | Low | DPST-9-4 | `7/3/121` Garage Helligkeit Tür BWM |
| 130 | Temperatur | `K....` | Low | DPST-9-1 | `7/1/121` Garage Temperatur Tür BWM |
| 131 | Taste links | `K....` | Low |  | `7/4/121` Garage-Präsenz-Innen-Externer-Taster |
| 132 | Taste links | `K....` | Low | DPST-3-7 | `1/2/124` Gruppe-Garage-Innen-LED-Band-Dimmen-Relativ |

### 🔌 1.1.102 - theLuxa P300 KNX - Garage Präsenz Aussen Tor (theLuxa P300 KNX)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Uhrzeit | `K....` | Low | DPST-10-1 | `0/0/1` Uhrzeit |
| 1 | Zeitanfrage | `K....` | Low | DPST-1-1 | `0/0/5` Datum / Uhrzeit anfordern |
| 2 | Helligkeitswert | `K....` | Low | DPST-9-4 | `7/3/125` Garage-Helligkeit-Aussen-Tor |
| 3 | Temperaturwert | `K....` | Low | DPST-9-1 | `7/1/1` Garage Aussen Temperatur |
| 22 | C2 Bewegung | `K....` | Low | DPST-1-1 | `7/4/126` Garage-Präsenz-Aussen-Tor-Links |
| 38 | C3 Bewegung | `K....` | Low | DPST-1-1 | `7/4/127` Garage-Präsenz-Aussen-Tor-Mitte |
| 54 | C4 Bewegung | `K....` | Low | DPST-1-1 | `7/4/128` Garage-Präsenz-Aussen-Tor-Rechts |

### 🔌 1.1.201 - Raumtemperatursensor 1-fach UP (Raumtemperatursensor 1-fach UP)
| Obj | Funktion / Name | Flags | Prio | DPT | Verknüpfte Gruppen |
|---|---|---|---|---|---|
| 0 | Temperaturmesswert | `K....` | Low | DPST-9-1 | `7/1/122` Garage Temperatur MDT Sensor |

## Gruppenadressen
| Adresse | Name | DPT | Hauptgruppe | Mittelgruppe |
|---|---|---|---|---|
| 0/0/1 | Uhrzeit | DPST-10-1 | Zentralfunktionen | Uhrzeit |
| 0/0/2 | Datum | DPST-11-1 | Zentralfunktionen | Uhrzeit |
| 0/0/3 | Uhrzeit und Datum | DPST-19-1 | Zentralfunktionen | Uhrzeit |
| 0/0/4 | Sommer- / Winterzeit | DPST-1-2 | Zentralfunktionen | Uhrzeit |
| 0/0/5 | Datum / Uhrzeit anfordern | DPST-1-17 | Zentralfunktionen | Uhrzeit |
| 0/0/6 | Astrofunktion: Tag = 0 / Nacht = 1 | DPST-1-24 | Zentralfunktionen | Uhrzeit |
| 0/0/7 | Astrofunktion Invertiert: Tag = 1 / Nacht = 0 | DPST-1-2 | Zentralfunktionen | Uhrzeit |
| 0/0/8 | Werktag (Mo - Fr) | DPST-1-2 | Zentralfunktionen | Uhrzeit |
| 0/0/100 | NTP Server Synch anfordern | DPST-1-17 | Zentralfunktionen | Uhrzeit |
| 0/1/0 | Gartentor-Auf | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/1 | Gartentor-Zu | - | Zentralfunktionen | Verschluss |
| 0/1/2 | Gartentor-Stop | - | Zentralfunktionen | Verschluss |
| 0/1/3 | Gartentor-Status-Bewegung | DPST-1-11 | Zentralfunktionen | Verschluss |
| 0/1/4 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/5 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/6 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/7 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/8 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/9 | /// --- Gartentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/10 | Garagentor-Fahren | DPST-1-8 | Zentralfunktionen | Verschluss |
| 0/1/11 | Garagentor-Stop | DPST-1-9 | Zentralfunktionen | Verschluss |
| 0/1/12 | Garagentor-Lüftung-Schmal | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/13 | Garagentor-Lüftung | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/14 | Garagentor-Status-Lüftung-Schmal | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/15 | Garagentor-Status-Lüftung | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/16 | Garagentor-Status-Position | DPST-5-1 | Zentralfunktionen | Verschluss |
| 0/1/17 | Garagentor-Status-Offen | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/18 | Garagentor-Status-Geschlossen | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/19 | Garagentor-Status-Bewegung-Öffnen | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/20 | Garagentor-Status-Bewegung-Schliessen | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/21 | Garagentor-Status-Vorwarnzeit-Aktiv | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/22 | Garagentor-Status-Aufhaltezeit-Aktiv | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/23 | Garagentor-Status-Inspektion-Notwendig | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/24 | Garagentor-Status-Fehler-Hohe-Prio | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/25 | Garagentor-Status-Fehler-Niedrige-Prio | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/26 | Garagentor-Status-Antrieb-Gesperrt | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/27 | Garagentor-Status-Zähler-Power-On | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/28 | Garagentor-Status-Zähler-Werks-Reset | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/29 | Garagentor-Status-Zähler-Torzyklen-Unvollst | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/30 | Garagentor-Status-Zähler-Torzyklen-Vollst | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/31 | Garagentor-Status-Zähler-Torzyklen-Unvollst-Nach-Wartung | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/32 | Garagentor-Status-Zähler-Fahrbefehle | DPST-12-1 | Zentralfunktionen | Verschluss |
| 0/1/33 | Garagentor-Status-Zähler-Laufzeit-Sekunden | DPST-13-100 | Zentralfunktionen | Verschluss |
| 0/1/34 | Garagentor-Status-Zähler-Laufzeit-Stunden | DPST-7-7 | Zentralfunktionen | Verschluss |
| 0/1/35 | Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Sekunden | DPST-13-100 | Zentralfunktionen | Verschluss |
| 0/1/36 | Garagentor-Status-Zähler-Laufzeit-Nach-Wartung-Stunden | DPST-7-7 | Zentralfunktionen | Verschluss |
| 0/1/37 | Garagentor-Status-Zähler-Betriebsstunden | DPST-13-100 | Zentralfunktionen | Verschluss |
| 0/1/38 | Garagentor-Status-Riegelstatus | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/39 | Garagentor-Status-Konfigurationsfehler | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/40 | Garagentor-Status-Torposition-Ungültig | DPST-1-2 | Zentralfunktionen | Verschluss |
| 0/1/41 | Garagentor-Status-Position-Invertiert | DPST-5-1 | Zentralfunktionen | Verschluss |
| 0/1/42 | /// --- Garage-Garagentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/43 | /// --- Garage-Garagentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/44 | /// --- Garage-Garagentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/45 | Alles Auf - Eingang | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/46 | Alles Auf - Ausgang 1 - Garagentor | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/47 | Alles Auf - Ausgang 2 - Gartentor | DPST-1-1 | Zentralfunktionen | Verschluss |
| 0/1/48 | /// --- Garage-Garagentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/49 | /// --- Garage-Garagentor-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/50 | Garage-Fenster-Auf/Ab | DPST-1-8 | Zentralfunktionen | Verschluss |
| 0/1/51 | Garage-Fenster-Stopp | DPST-1-17 | Zentralfunktionen | Verschluss |
| 0/1/52 | Garage-Fenster-Status-Bewegung | DPST-1-11 | Zentralfunktionen | Verschluss |
| 0/1/53 | Garage-Fenster-Status-Absolute-Position | DPST-5-1 | Zentralfunktionen | Verschluss |
| 0/1/54 | Garage-Fenster-Status-Diagnosetext | DPST-16-0 | Zentralfunktionen | Verschluss |
| 0/1/55 | Garage-Fenster-Status-aktuelle-Richtung | DPST-1-8 | Zentralfunktionen | Verschluss |
| 0/1/56 | Garage-Fenster-Status-Obere-Position | DPST-1-11 | Zentralfunktionen | Verschluss |
| 0/1/57 | Garage-Fenster-Status-Untere-Position | DPST-1-11 | Zentralfunktionen | Verschluss |
| 0/1/58 | /// --- Garage-Fenster-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/1/59 | /// --- Garage-Fenster-Reserve --- /// | - | Zentralfunktionen | Verschluss |
| 0/2/0 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/1 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/2 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/3 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/4 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/5 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/6 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/7 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/8 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/2/9 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Logik |
| 0/4/0 | /// --- Reserve --- /// | - | Zentralfunktionen | Licht |
| 0/4/1 | Licht-Zentral | DPST-1-1 | Zentralfunktionen | Licht |
| 0/4/2 | Licht-Zentral-Status | DPST-1-11 | Zentralfunktionen | Licht |
| 0/4/10 | Licht-Zentral-Innen | DPST-1-1 | Zentralfunktionen | Licht |
| 0/4/11 | Licht-Zentral-Innen-Status | DPST-1-11 | Zentralfunktionen | Licht |
| 0/4/12 | Licht-Zentral-Innen-Unten | DPST-1-1 | Zentralfunktionen | Licht |
| 0/4/13 | Licht-Zentral-Innen-Unten-Status | DPST-1-11 | Zentralfunktionen | Licht |
| 0/4/14 | Licht-Zentral-Innen-Oben | DPST-1-1 | Zentralfunktionen | Licht |
| 0/4/15 | Licht-Zentral-Innen-Oben-Status | DPST-1-11 | Zentralfunktionen | Licht |
| 0/4/20 | Licht-Zentral-Aussen | DPST-1-1 | Zentralfunktionen | Licht |
| 0/4/21 | Licht-Zentral-Aussen-Status | DPST-1-11 | Zentralfunktionen | Licht |
| 1/0/1 | EG-Flur-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/2 | EG-Flur-Spots-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/3 | Flur-Treppenlicht-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/4 | Gruppe-Flur-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/5 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/6 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/7 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/8 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/9 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/10 | EG-Gaestebad-Spiegel-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/11 | EG-Gaestebad-Fenster-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/12 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/13 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/14 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/15 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/16 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/17 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/18 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/19 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/20 | EG-HWR-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/21 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/22 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/23 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/24 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/25 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/26 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/27 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/28 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/29 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/30 | EG-WZ-Couch-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/31 | EG-WZ-Wandleuchten-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/32 | EG-WZ-Steckdosen-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/33 | Gruppe-WZ-EZ-Steckdosen-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/34 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/35 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/36 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/37 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/38 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/39 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/40 | EG-Esszimmer-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/41 | EG-Esszimmer-Steckdosen-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/42 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/43 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/44 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/45 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/46 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/47 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/48 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/49 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/50 | EG-Kueche-Fenster-Spots-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/51 | EG-Kueche-Vorn-Deckenleuchten-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/52 | EG-Kueche-Hinten-Deckenleuchten-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/53 | Gruppe-Kueche-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/54 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/55 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/56 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/57 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/58 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/59 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/60 | OG-Flur-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/61 | OG-Flur-Wandleuchten-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/62 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/63 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/64 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/65 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/66 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/67 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/68 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/69 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/70 | OG-Gaestezimmer-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/71 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/72 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/73 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/74 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/75 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/76 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/77 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/78 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/79 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/80 | OG-Schlafzimmer-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/81 | OG-Schlafzimmer-Ankleide-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/82 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/83 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/84 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/85 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/86 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/87 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/88 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/89 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/90 | OG-Arbeitszimmer-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/91 | OG-Arbeitszimmer-Schreibtisch-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/92 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/93 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/94 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/95 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/96 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/97 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/98 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/99 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/100 | OG-Badezimmer-Spiegel-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/101 | OG-Badezimmer-Deckenleuchte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/102 | OG-Badezimmer-Spots-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/103 | Gruppe-OG-Badezimmer-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/104 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/105 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/106 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/107 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/108 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/109 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/110 | Haus-Aussenbeleuchtung-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/111 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/112 | Haustuer-Aussenbeleuchtung-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/113 | Garten-Weg-Vorn-Licht-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/114 | Garten-Beet-Hinten-Licht-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/115 | Gruppe-Garten-Wegleuchten-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/116 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/117 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/118 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/119 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/120 | Garage-Innen-LED-Band-Nachbar-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/121 | Garage-Innen-LED-Band-Tuer-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/122 | Garage-Innen-LED-Band-Feld-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/123 | Garage-Innen-LED-Band-Mitte-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/124 | Gruppe-Garage-Innen-LED-Band-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/125 | Garage-Aussen-Downlights-Tor-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/126 | Garage-Aussen-Downlights-Tuer-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/127 | Gruppe-Garage-Aussen-Downlights-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/128 | Garage-Aussen-Steckdosen-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/129 | Garagentor-Licht-Schalten | DPST-1-1 | Beleuchtung / Standard | Schalten |
| 1/0/130 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/131 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/132 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/133 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/134 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/135 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/136 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/137 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/138 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/139 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/140 | Gartenhaus-Innen-Links-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/141 | Gartenhaus-Innen-Rechts-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/142 | Gartenhaus-Aussen-Schalten | - | Beleuchtung / Standard | Schalten |
| 1/0/143 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/144 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/145 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/146 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/147 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/148 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/149 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/150 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/151 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/152 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/153 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/154 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/155 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/156 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/157 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/158 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/0/159 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Schalten |
| 1/1/1 | EG-Flur-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/2 | EG-Flur-Spots-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/3 | Flur-Treppenlicht-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/4 | Gruppe-Flur-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/5 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/6 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/7 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/8 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/9 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/10 | EG-Gaestebad-Spiegel-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/11 | EG-Gaestebad-Fenster-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/12 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/13 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/14 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/15 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/16 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/17 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/18 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/19 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/20 | EG-HWR-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/21 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/22 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/23 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/24 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/25 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/26 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/27 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/28 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/29 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/30 | EG-WZ-Couch-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/31 | EG-WZ-Wandleuchten-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/32 | EG-WZ-Steckdosen-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/33 | Gruppe-WZ-EZ-Steckdosen-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/34 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/35 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/36 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/37 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/38 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/39 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/40 | EG-Esszimmer-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/41 | EG-Esszimmer-Steckdosen-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/42 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/43 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/44 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/45 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/46 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/47 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/48 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/49 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/50 | EG-Kueche-Fenster-Spots-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/51 | EG-Kueche-Vorn-Deckenleuchten-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/52 | EG-Kueche-Hinten-Deckenleuchten-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/53 | Gruppe-Kueche-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/54 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/55 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/56 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/57 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/58 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/59 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/60 | OG-Flur-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/61 | OG-Flur-Wandleuchten-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/62 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/63 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/64 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/65 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/66 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/67 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/68 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/69 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/70 | OG-Gaestezimmer-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/71 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/72 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/73 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/74 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/75 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/76 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/77 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/78 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/79 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/80 | OG-Schlafzimmer-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/81 | OG-Schlafzimmer-Ankleide-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/82 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/83 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/84 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/85 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/86 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/87 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/88 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/89 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/90 | OG-Arbeitszimmer-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/91 | OG-Arbeitszimmer-Schreibtisch-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/92 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/93 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/94 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/95 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/96 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/97 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/98 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/99 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/100 | OG-Badezimmer-Spiegel-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/101 | OG-Badezimmer-Deckenleuchte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/102 | OG-Badezimmer-Spots-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/103 | Gruppe-OG-Badezimmer-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/104 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/105 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/106 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/107 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/108 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/109 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/110 | Haus-Aussenbeleuchtung-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/111 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/112 | Haustuer-Aussenbeleuchtung-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/113 | Garten-Weg-Vorn-Licht-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/114 | Garten-Beet-Hinten-Licht-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/115 | Gruppe-Garten-Wegleuchten-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/116 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/117 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/118 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/119 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/120 | Garage-Innen-LED-Band-Nachbar-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/121 | Garage-Innen-LED-Band-Tuer-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/122 | Garage-Innen-LED-Band-Feld-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/123 | Garage-Innen-LED-Band-Mitte-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/124 | Gruppe-Garage-Innen-LED-Band-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/125 | Garage-Aussen-Downlights-Tor-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/126 | Garage-Aussen-Downlights-Tuer-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/127 | Gruppe-Garage-Aussen-Downlights-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/128 | Garage-Aussen-Steckdosen-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/129 | Garagentor-Licht-Status-Schalten | DPST-1-11 | Beleuchtung / Standard | Status Schalten |
| 1/1/130 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/131 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/132 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/133 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/134 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/135 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/136 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/137 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/138 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/139 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/140 | Gartenhaus-Innen-Links-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/141 | Gartenhaus-Innen-Rechts-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/142 | Gartenhaus-Aussen-Status-Schalten | - | Beleuchtung / Standard | Status Schalten |
| 1/1/143 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/144 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/145 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/146 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/147 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/148 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/149 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/150 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/151 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/152 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/153 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/154 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/155 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/156 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/157 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/158 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/1/159 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Schalten |
| 1/2/1 | EG-Flur-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/2 | EG-Flur-Spots-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/3 | Flur-Treppenlicht-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/4 | Gruppe-Flur-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/5 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/6 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/7 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/8 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/9 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/10 | EG-Gaestebad-Spiegel-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/11 | EG-Gaestebad-Fenster-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/12 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/13 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/14 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/15 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/16 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/17 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/18 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/19 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/20 | EG-HWR-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/21 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/22 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/23 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/24 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/25 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/26 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/27 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/28 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/29 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/30 | EG-WZ-Couch-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/31 | EG-WZ-Wandleuchten-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/32 | EG-WZ-Steckdosen-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/33 | Gruppe-WZ-EZ-Steckdosen-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/34 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/35 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/36 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/37 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/38 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/39 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/40 | EG-Esszimmer-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/41 | EG-Esszimmer-Steckdosen-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/42 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/43 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/44 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/45 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/46 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/47 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/48 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/49 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/50 | EG-Kueche-Fenster-Spots-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/51 | EG-Kueche-Vorn-Deckenleuchten-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/52 | EG-Kueche-Hinten-Deckenleuchten-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/53 | Gruppe-Kueche-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/54 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/55 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/56 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/57 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/58 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/59 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/60 | OG-Flur-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/61 | OG-Flur-Wandleuchten-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/62 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/63 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/64 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/65 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/66 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/67 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/68 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/69 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/70 | OG-Gaestezimmer-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/71 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/72 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/73 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/74 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/75 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/76 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/77 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/78 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/79 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/80 | OG-Schlafzimmer-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/81 | OG-Schlafzimmer-Ankleide-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/82 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/83 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/84 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/85 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/86 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/87 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/88 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/89 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/90 | OG-Arbeitszimmer-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/91 | OG-Arbeitszimmer-Schreibtisch-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/92 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/93 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/94 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/95 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/96 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/97 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/98 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/99 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/100 | OG-Badezimmer-Spiegel-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/101 | OG-Badezimmer-Deckenleuchte-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/102 | OG-Badezimmer-Spots-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/103 | Gruppe-OG-Badezimmer-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/104 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/105 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/106 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/107 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/108 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/109 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/110 | Haus-Aussenbeleuchtung-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/111 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/112 | Haustuer-Aussenbeleuchtung-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/113 | Garten-Weg-Vorn-Licht-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/114 | Garten-Beet-Hinten-Licht-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/115 | Gruppe-Garten-Wegleuchten-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/116 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/117 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/118 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/119 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/120 | Garage-Innen-LED-Band-Nachbar-Dimmen-Relativ | DPST-3-7 | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/121 | Garage-Innen-LED-Band-Tuer-Dimmen-Relativ | DPST-3-7 | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/122 | Garage-Innen-LED-Band-Feld-Dimmen-Relativ | DPST-3-7 | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/123 | Garage-Innen-LED-Band-Mitte-Dimmen-Relativ | DPST-3-7 | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/124 | Gruppe-Garage-Innen-LED-Band-Dimmen-Relativ | DPST-3-7 | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/125 | Garage-Aussen-Downlights-Tor-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/126 | Garage-Aussen-Downlights-Tuer-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/127 | Gruppe-Garage-Aussen-Downlights-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/128 | Garage-Aussen-Steckdosen-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/129 | Garagentor-Licht-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/130 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/131 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/132 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/133 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/134 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/135 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/136 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/137 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/138 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/139 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/140 | Gartenhaus-Innen-Links-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/141 | Gartenhaus-Innen-Rechts-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/142 | Gartenhaus-Aussen-Dimmen-Relativ | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/143 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/144 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/145 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/146 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/147 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/148 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/149 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/150 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/151 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/152 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/153 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/154 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/155 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/156 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/157 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/158 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/2/159 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Relativ |
| 1/3/1 | EG-Flur-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/2 | EG-Flur-Spots-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/3 | Flur-Treppenlicht-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/4 | Gruppe-Flur-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/5 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/6 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/7 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/8 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/9 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/10 | EG-Gaestebad-Spiegel-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/11 | EG-Gaestebad-Fenster-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/12 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/13 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/14 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/15 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/16 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/17 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/18 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/19 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/20 | EG-HWR-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/21 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/22 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/23 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/24 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/25 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/26 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/27 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/28 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/29 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/30 | EG-WZ-Couch-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/31 | EG-WZ-Wandleuchten-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/32 | EG-WZ-Steckdosen-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/33 | Gruppe-WZ-EZ-Steckdosen-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/34 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/35 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/36 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/37 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/38 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/39 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/40 | EG-Esszimmer-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/41 | EG-Esszimmer-Steckdosen-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/42 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/43 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/44 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/45 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/46 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/47 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/48 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/49 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/50 | EG-Kueche-Fenster-Spots-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/51 | EG-Kueche-Vorn-Deckenleuchten-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/52 | EG-Kueche-Hinten-Deckenleuchten-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/53 | Gruppe-Kueche-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/54 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/55 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/56 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/57 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/58 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/59 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/60 | OG-Flur-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/61 | OG-Flur-Wandleuchten-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/62 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/63 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/64 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/65 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/66 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/67 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/68 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/69 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/70 | OG-Gaestezimmer-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/71 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/72 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/73 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/74 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/75 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/76 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/77 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/78 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/79 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/80 | OG-Schlafzimmer-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/81 | OG-Schlafzimmer-Ankleide-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/82 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/83 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/84 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/85 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/86 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/87 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/88 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/89 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/90 | OG-Arbeitszimmer-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/91 | OG-Arbeitszimmer-Schreibtisch-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/92 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/93 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/94 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/95 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/96 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/97 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/98 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/99 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/100 | OG-Badezimmer-Spiegel-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/101 | OG-Badezimmer-Deckenleuchte-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/102 | OG-Badezimmer-Spots-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/103 | Gruppe-OG-Badezimmer-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/104 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/105 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/106 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/107 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/108 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/109 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/110 | Haus-Aussenbeleuchtung-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/111 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/112 | Haustuer-Aussenbeleuchtung-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/113 | Garten-Weg-Vorn-Licht-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/114 | Garten-Beet-Hinten-Licht-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/115 | Gruppe-Garten-Wegleuchten-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/116 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/117 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/118 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/119 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/120 | Garage-Innen-LED-Band-Nachbar-Dimmen-Absolut | DPST-5-1 | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/121 | Garage-Innen-LED-Band-Tuer-Dimmen-Absolut | DPST-5-1 | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/122 | Garage-Innen-LED-Band-Feld-Dimmen-Absolut | DPST-5-1 | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/123 | Garage-Innen-LED-Band-Mitte-Dimmen-Absolut | DPST-5-1 | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/124 | Gruppe-Garage-Innen-LED-Band-Dimmen-Absolut | DPST-5-1 | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/125 | Garage-Aussen-Downlights-Tor-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/126 | Garage-Aussen-Downlights-Tuer-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/127 | Gruppe-Garage-Aussen-Downlights-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/128 | Garage-Aussen-Steckdosen-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/129 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/130 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/131 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/132 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/133 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/134 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/135 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/136 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/137 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/138 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/139 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/140 | Gartenhaus-Innen-Links-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/141 | Gartenhaus-Innen-Rechts-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/142 | Gartenhaus-Aussen-Dimmen-Absolut | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/143 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/144 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/145 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/146 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/147 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/148 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/149 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/150 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/151 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/152 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/153 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/154 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/155 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/156 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/157 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/158 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/3/159 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Dimmen Absolut |
| 1/4/1 | EG-Flur-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/2 | EG-Flur-Spots-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/3 | Flur-Treppenlicht-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/4 | Gruppe-Flur-Status-Dimmwert | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/5 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/6 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/7 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/8 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/9 | /// --- EG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/10 | EG-Gaestebad-Spiegel-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/11 | EG-Gaestebad-Fenster-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/12 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/13 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/14 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/15 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/16 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/17 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/18 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/19 | /// --- EG-Gaestebad-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/20 | EG-HWR-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/21 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/22 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/23 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/24 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/25 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/26 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/27 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/28 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/29 | /// --- EG-HWR-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/30 | EG-WZ-Couch-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/31 | EG-WZ-Wandleuchten-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/32 | EG-WZ-Steckdosen-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/33 | Gruppe-WZ-EZ-Steckdosen-Status-Dimmwert | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/34 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/35 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/36 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/37 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/38 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/39 | /// --- EG-WZ-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/40 | EG-Esszimmer-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/41 | EG-Esszimmer-Steckdosen-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/42 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/43 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/44 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/45 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/46 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/47 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/48 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/49 | /// --- EG-Esszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/50 | EG-Kueche-Fenster-Spots-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/51 | EG-Kueche-Vorn-Deckenleuchten-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/52 | EG-Kueche-Hinten-Deckenleuchten-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/53 | Gruppe-Kueche-Schalten-Status-Dimmwert | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/54 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/55 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/56 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/57 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/58 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/59 | /// --- EG-Kueche-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/60 | OG-Flur-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/61 | OG-Flur-Wandleuchten-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/62 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/63 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/64 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/65 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/66 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/67 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/68 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/69 | /// --- OG-Flur-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/70 | OG-Gaestezimmer-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/71 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/72 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/73 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/74 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/75 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/76 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/77 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/78 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/79 | /// --- OG-Gaestezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/80 | OG-Schlafzimmer-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/81 | OG-Schlafzimmer-Ankleide-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/82 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/83 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/84 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/85 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/86 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/87 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/88 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/89 | /// --- OG-Schlafzimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/90 | OG-Arbeitszimmer-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/91 | OG-Arbeitszimmer-Schreibtisch-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/92 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/93 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/94 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/95 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/96 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/97 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/98 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/99 | /// --- OG-Arbeitszimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/100 | OG-Badezimmer-Spiegel-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/101 | OG-Badezimmer-Deckenleuchte-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/102 | OG-Badezimmer-Spots-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/103 | Gruppe-OG-Badezimmer-Status-Dimmwert | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/104 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/105 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/106 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/107 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/108 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/109 | /// --- OG-Badezimmer-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/110 | Haus-Aussenbeleuchtung-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/111 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/112 | Haustuer-Aussenbeleuchtung-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/113 | Garten-Weg-Vorn-Licht-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/114 | Garten-Beet-Hinten-Licht-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/115 | Gruppe-Garten-Wegleuchten-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/116 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/117 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/118 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/119 | /// --- Garten-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/120 | Garage-Innen-LED-Band-Nachbar-Status-Dimmwerrt | DPST-5-1 | Beleuchtung / Standard | Status Dimmwert |
| 1/4/121 | Garage-Innen-LED-Band-Tuer-Status-Dimmwerrt | DPST-5-1 | Beleuchtung / Standard | Status Dimmwert |
| 1/4/122 | Garage-Innen-LED-Band-Feld-Status-Dimmwerrt | DPST-5-1 | Beleuchtung / Standard | Status Dimmwert |
| 1/4/123 | Garage-Innen-LED-Band-Mitte-Status-Dimmwerrt | DPST-5-1 | Beleuchtung / Standard | Status Dimmwert |
| 1/4/124 | Gruppe-Garage-Innen-LED-Band-Status-Dimmwerrt | DPST-5-1 | Beleuchtung / Standard | Status Dimmwert |
| 1/4/125 | Garage-Aussen-Downlights-Tor-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/126 | Garage-Aussen-Downlights-Tuer-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/127 | Gruppe-Garage-Aussen-Downlights-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/128 | Garage-Aussen-Steckdosen-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/129 | Garagentor-Licht-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/130 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/131 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/132 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/133 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/134 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/135 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/136 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/137 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/138 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/139 | /// --- Garage-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/140 | Gartenhaus-Innen-Links-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/141 | Gartenhaus-Innen-Rechts-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/142 | Gartenhaus-Aussen-Status-Dimmwerrt | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/143 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/144 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/145 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/146 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/147 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/148 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/149 | /// --- Gartenhaus-Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/150 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/151 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/152 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/153 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/154 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/155 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/156 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/157 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/158 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/4/159 | /// --- Reserve --- /// | - | Beleuchtung / Standard | Status Dimmwert |
| 1/7/120 | Garage-LED-Band-Dimmer-Überstrom-Alarm | DPST-1-5 | Beleuchtung / Standard | Sonstiges |
| 1/7/121 | Garage-LED-Band-Dimmer-Übertemperatur-Alarm | DPST-1-5 | Beleuchtung / Standard | Sonstiges |
| 1/7/122 | Garage-LED-Band-Dimmer-Spannungsversorgung | DPST-1-11 | Beleuchtung / Standard | Sonstiges |
| 7/0/0 | Haus-Netzteil-Anzahl-Spannungsausfälle | DPST-7-1 | Sensoren | Zentral |
| 7/0/1 | Haus-Netzteil-LetzterSpannungsausfall-Uhrzeit | DPST-10-1 | Sensoren | Zentral |
| 7/0/2 | Haus-Netzteil-LetzterSpannungsausfall-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/3 | Haus-Netzteil-Anzahl-Busresets-Bus-A | DPST-7-1 | Sensoren | Zentral |
| 7/0/4 | Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-A | DPST-10-1 | Sensoren | Zentral |
| 7/0/5 | Haus-Netzteil-Letzter-Busreset-Datum-Bus-A | DPST-11-1 | Sensoren | Zentral |
| 7/0/6 | Haus-Netzteil-Anzahl-Busresets-Bus-B | DPST-7-1 | Sensoren | Zentral |
| 7/0/7 | Haus-Netzteil-Letzter-Busreset-Uhrzeit-Bus-B | DPST-10-1 | Sensoren | Zentral |
| 7/0/8 | Haus-Netzteil-Letzter-Busreset-Datum-Bus-B | DPST-11-1 | Sensoren | Zentral |
| 7/0/9 | Haus-Netzteil-AnzahlBusneustarts-Bus-A | DPST-7-1 | Sensoren | Zentral |
| 7/0/10 | Haus-Netzteil-Letzter-Busneustart-Uhrzeit-Bus-A | DPST-10-1 | Sensoren | Zentral |
| 7/0/11 | Haus-Netzteil-Letzter-Busneustart-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/12 | Remote-Bus-Reset-Bus-A | DPST-1-17 | Sensoren | Zentral |
| 7/0/13 | Remote-Bus-Reset-Bus-B | DPST-1-17 | Sensoren | Zentral |
| 7/0/14 | Haus-Netzteil-Analysereset | DPST-1-17 | Sensoren | Zentral |
| 7/0/16 | Haus-Netzteil-Betreibsstunden-Lebenszeit | DPST-7-7 | Sensoren | Zentral |
| 7/0/17 | Haus-Netzteil-Betriebssekunden-Lebenszeit | DPST-13-100 | Sensoren | Zentral |
| 7/0/20 | Haus-Netzteil-Spannung-Bus-A | DPST-14-27 | Sensoren | Zentral |
| 7/0/21 | Haus-Netzteil-Spannung-Bus-B | DPST-14-27 | Sensoren | Zentral |
| 7/0/22 | Haus-Netzteil-Spannung-Aux | DPST-14-27 | Sensoren | Zentral |
| 7/0/23 | Haus-Netzteil-Strom-Bus-A | DPST-9-21 | Sensoren | Zentral |
| 7/0/24 | Haus-Netzteil-Strom-Bus-B | DPST-9-21 | Sensoren | Zentral |
| 7/0/25 | Haus-Netzteil-Strom-Aux | DPST-9-21 | Sensoren | Zentral |
| 7/0/26 | Haus-Netzteil-Strom-Gesamt | DPST-9-21 | Sensoren | Zentral |
| 7/0/27 | Haus-Netzteil-Leistung-Bus-A | DPST-14-56 | Sensoren | Zentral |
| 7/0/28 | Haus-Netzteil-Leistung-Bus-B | DPST-14-56 | Sensoren | Zentral |
| 7/0/29 | Haus-Netzteil-Leistung-Aux | DPST-14-56 | Sensoren | Zentral |
| 7/0/30 | Haus-Netzteil-Leistung-Gesamt | DPST-14-56 | Sensoren | Zentral |
| 7/0/31 | Haus-Netzteil-Temperatur | DPST-9-1 | Sensoren | Zentral |
| 7/0/32 | Haus-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde-Bus-A | DPST-7-1 | Sensoren | Zentral |
| 7/0/33 | Haus-Netzteil-Mittlere-Telegrammrate-pro-Sekunde-Bus-B | DPST-7-1 | Sensoren | Zentral |
| 7/0/34 | Haus-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit | DPST-13-10 | Sensoren | Zentral |
| 7/0/35 | Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt | DPST-13-10 | Sensoren | Zentral |
| 7/0/36 | Haus-Netzteil-Abgegebene-Gesamtenergie-seit-Reset | DPST-13-10 | Sensoren | Zentral |
| 7/0/37 | Haus-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit | DPST-13-10 | Sensoren | Zentral |
| 7/0/38 | Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt | DPST-13-10 | Sensoren | Zentral |
| 7/0/39 | Haus-Netzteil-Aufgenommene-Gesamtenergie-seit-Reset | DPST-13-10 | Sensoren | Zentral |
| 7/0/40 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/41 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/42 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/43 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/44 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/45 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/46 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/47 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/48 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/49 | /// --- Haus-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/50 | Garage-Netzteil-Anzahl-Spannungsausfälle | DPST-7-1 | Sensoren | Zentral |
| 7/0/51 | Garage-Netzteil-LetzterSpannungsausfall-Uhrzeit | DPST-10-1 | Sensoren | Zentral |
| 7/0/52 | Garage-Netzteil-LetzterSpannungsausfall-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/53 | Garage-Netzteil-Anzahl-Busresets | DPST-7-1 | Sensoren | Zentral |
| 7/0/54 | Garage-Netzteil-Letzter-Busreset-Uhrzeit | DPST-10-1 | Sensoren | Zentral |
| 7/0/55 | Garage-Netzteil-Letzter-Busreset-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/56 | Garage-Netzteil-Anzahlspannungsreset-AuxA | DPST-7-1 | Sensoren | Zentral |
| 7/0/57 | Garage-Netzteil-Anzahlspannungsreset-AuxA-Uhrzeit | DPST-10-1 | Sensoren | Zentral |
| 7/0/58 | Garage-Netzteil-Anzahlspannungsreset-AuxA-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/59 | Garage-Netzteil-Anzahl-Busneustarts | DPST-7-1 | Sensoren | Zentral |
| 7/0/60 | Garage-Netzteil-Letzter-Busneustart-Uhrzeit | DPST-10-1 | Sensoren | Zentral |
| 7/0/61 | Garage-Netzteil-Letzter-Busneustart-Datum | DPST-11-1 | Sensoren | Zentral |
| 7/0/62 | Garage-Netzteil-Analysereset | DPST-1-17 | Sensoren | Zentral |
| 7/0/63 | Garage-Netzteil-Status | DPST-16-1 | Sensoren | Zentral |
| 7/0/64 | Garage-Netzteil-Betriebsstunden-Lebenszeit | DPST-7-7 | Sensoren | Zentral |
| 7/0/65 | Garage-Netzteil-Betriebssekunden-Lebenszeit | DPST-13-100 | Sensoren | Zentral |
| 7/0/66 | Garage-Netzteil-RemoteBusreset | DPST-1-17 | Sensoren | Zentral |
| 7/0/67 | Garage-Netzteil-Spannungsreset-AuxA | DPST-1-17 | Sensoren | Zentral |
| 7/0/68 | Garage-Netzteil-Spannung-Bus | DPST-14-27 | Sensoren | Zentral |
| 7/0/69 | Garage-Netzteil-Spannung-AuxA | DPST-14-27 | Sensoren | Zentral |
| 7/0/70 | Garage-Netzteil-Spannung-AuxB | DPST-14-27 | Sensoren | Zentral |
| 7/0/71 | Garage-Netzteil-Strom-Bus | DPST-9-21 | Sensoren | Zentral |
| 7/0/72 | Garage-Netzteil-Strom-AuxA | DPST-9-21 | Sensoren | Zentral |
| 7/0/73 | Garage-Netzteil-Strom-AuxB | DPST-9-21 | Sensoren | Zentral |
| 7/0/74 | Garage-Netzteil-Strom-Gesamt | DPST-9-21 | Sensoren | Zentral |
| 7/0/75 | Garage-Netzteil-Leistung-Bus | DPST-14-56 | Sensoren | Zentral |
| 7/0/76 | Garage-Netzteil-Leistung-AuxA | DPST-14-56 | Sensoren | Zentral |
| 7/0/77 | Garage-Netzteil-Leistung-AuxB | DPST-14-56 | Sensoren | Zentral |
| 7/0/78 | Garage-Netzteil-Leistung-Gesamt | DPST-14-56 | Sensoren | Zentral |
| 7/0/79 | Garage-Netzteil-Temperatur | DPST-9-1 | Sensoren | Zentral |
| 7/0/80 | Garage-Netzteil-Aktuelle-Telegrammrate-pro-Sekunde | DPST-7-1 | Sensoren | Zentral |
| 7/0/81 | Garage-Netzteil-Mittlere-Telegrammrate-pro-Sekunde | DPST-7-1 | Sensoren | Zentral |
| 7/0/82 | Garage-Netzteil-Abgegebene-Gesamtenergie-Lebenszeit | DPST-13-10 | Sensoren | Zentral |
| 7/0/83 | Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Einschaltzeitpunkt | DPST-13-10 | Sensoren | Zentral |
| 7/0/84 | Garage-Netzteil-Abgegebene-Gesamtenergie-seit-Analysereset | DPST-13-10 | Sensoren | Zentral |
| 7/0/85 | Garage-Netzteil-Aufgenommene-Gesamtenergie-Lebenszeit | DPST-13-10 | Sensoren | Zentral |
| 7/0/86 | Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Einschaltzeitpunkt | DPST-13-10 | Sensoren | Zentral |
| 7/0/87 | Garage-Netzteil-Aufgenommene-Gesamtenergie-seit-Analysereset | DPST-13-10 | Sensoren | Zentral |
| 7/0/88 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/89 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/90 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/91 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/92 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/93 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/94 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/95 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/96 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/97 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/98 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/0/99 | /// --- Garage-Netzteil-Reserve --- /// | - | Sensoren | Zentral |
| 7/1/1 | Garage Aussen Temperatur | DPST-9-1 | Sensoren | Temperatur |
| 7/1/50 | Gartenhaus Temperatur  | DPST-9-1 | Sensoren | Temperatur |
| 7/1/120 | Garage Temperatur Decke BWM | DPT-9 | Sensoren | Temperatur |
| 7/1/121 | Garage Temperatur Tür BWM | DPST-9-1 | Sensoren | Temperatur |
| 7/1/122 | Garage Temperatur MDT Sensor | DPST-9-1 | Sensoren | Temperatur |
| 7/2/1 | Regenalarm | DPST-1-1 | Sensoren | Feuchtigkeit |
| 7/2/50 | Gartenhaus Feuchtigkeit | DPST-9-7 | Sensoren | Feuchtigkeit |
| 7/3/120 | Garage Helligkeit Decken BWM | DPT-9 | Sensoren | Helligkeit |
| 7/3/121 | Garage Helligkeit Tür BWM | DPST-9-4 | Sensoren | Helligkeit |
| 7/3/125 | Garage-Helligkeit-Aussen-Tor | DPST-9-4 | Sensoren | Helligkeit |
| 7/4/120 | Garage-Präsenz-Innen-Externe-Bewegung | DPST-1-1 | Sensoren | Präsenz |
| 7/4/121 | Garage-Präsenz-Innen-Externer-Taster | DPST-1-1 | Sensoren | Präsenz |
| 7/4/124 | Garage-Präsenz-Innen-Meldung | DPST-1-1 | Sensoren | Präsenz |
| 7/4/125 | Garage-Präsenz-Innen-Meldung-Extern | DPST-1-1 | Sensoren | Präsenz |
| 7/4/126 | Garage-Präsenz-Aussen-Tor-Links | DPST-1-1 | Sensoren | Präsenz |
| 7/4/127 | Garage-Präsenz-Aussen-Tor-Mitte | DPST-1-1 | Sensoren | Präsenz |
| 7/4/128 | Garage-Präsenz-Aussen-Tor-Rechts | DPST-1-1 | Sensoren | Präsenz |
