import numpy as np
import matplotlib.pyplot as plt

# Gegebene Werte
body_weight = 70  # kg
Vd = 0.7 * body_weight  # Verteilungsvolumen in L
CrCl = 30 / 1000 * 60  # Nierenfunktion Clearance in L/h (30 mL/min)
kel = CrCl / Vd  # Eliminationskonstante in h^-1
t_half = np.log(2) / kel  # Halbwertszeit in Stunden
dose_loading = 1500  # Lade-Dosis in mg
dose_maintenance = 1250  # Erhaltungsdosis in mg
t_inf = 3  # Infusionsdauer in Stunden
t_dosage_interval = 24  # Dosierungsintervall in Stunden
treatment_duration = 7 * 24  # Behandlungsdauer in Stunden
dt = 0.5  # Zeitschritt in Stunden

# Zeitpunkte und initiale Konzentration
time_points = np.arange(0, treatment_duration + dt, dt)  # in Stunden
plasma_concentration = np.zeros_like(time_points, dtype=float)

# Initialisierung: Lade-Dosis
plasma_concentration[0] = dose_loading / Vd  # Initialkonzentration nach Lade-Dosis
A = dose_loading  # Anfangsmenge in mg

# Liste zur Verfolgung der Endzeiten laufender Infusionen
infusions = []

for i in range(1, len(time_points)):
    t = time_points[i]

    # Entfernen abgeschlossener Infusionen
    infusions = [end_time for end_time in infusions if t < end_time]

    # Berechnung der aktuellen Infusionsrate (mg/h)
    infusion_rate = (dose_maintenance / t_inf) * len(infusions)

    # Berechnung der Elimination und Infusion
    # Formel für kontinuierliche Infusion unter Berücksichtigung der Elimination:
    # A(t) = A(t-Δt) * exp(-kel * Δt) + (Infusionsrate / kel) * (1 - exp(-kel * Δt))
    A = A * np.exp(-kel * dt) + (infusion_rate / kel) * (1 - np.exp(-kel * dt))

    # Aktualisierung der Plasmakonzentration
    plasma_concentration[i] = A / Vd

    # Überprüfung, ob eine neue Erhaltungsdosis verabreicht wird
    # Verwendung von np.isclose zur Berücksichtigung von Rundungsfehlern
    if np.isclose(t % t_dosage_interval, 0, atol=dt / 2):
        infusions.append(t + t_inf)  # Hinzufügen der Endzeit der neuen Infusion

# Ergebnis plotten
plt.figure(figsize=(12, 6))
plt.plot(time_points, plasma_concentration, label="Vancomycin-Plasmaspiegel", color='b')
plt.xlabel("Zeit (Stunden)")
plt.ylabel("Plasmaspiegel (mg/L)")
plt.title("Pharmakokinetische Simulation der Vancomycin-Plasmaspiegel über 7 Tage")
plt.axhline(y=15, color='r', linestyle='--', label='Therapeutischer Bereich (15-20 mg/L)')
plt.axhline(y=20, color='r', linestyle='--')
plt.legend()
plt.grid()
plt.show()
