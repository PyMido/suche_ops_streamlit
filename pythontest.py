import numpy as np
import matplotlib.pyplot as plt

# Gegebene Werte
body_weight = 70  # kg
Vd = 0.7 * body_weight  # Verteilungsvolumen in L
CrCl = 30 / 1000 * 60  # Nierenfunktion Clearance in L/h
kel = CrCl / Vd  # Eliminationskonstante in h^-1
t_half = np.log(2) / kel  # Halbwertszeit in Stunden
dose_loading = 1500  # Lade-Dosis in mg
dose_maintenance = 1000  # Erhaltungsdosis in mg
t_inf = 20  # Infusionsdauer in Stunden
t_dosage_interval = 24  # Dosierungsintervall in Stunden
treatment_duration = 7 * 24  # Behandlungsdauer in Stunden

# Zeitpunkte und initiale Konzentration
time_points = np.arange(0, treatment_duration + t_dosage_interval, 0.5)  # in Stunden
plasma_concentration = np.zeros_like(time_points, dtype=float)

# Lade-Dosis initiale Konzentration
plasma_concentration[0] = dose_loading / Vd  # Initialkonzentration nach Lade-Dosis

# Pharmakokinetische Simulation für 10 Tage
for i, t in enumerate(time_points[1:], 1):
    previous_concentration = plasma_concentration[i - 1] * np.exp(-kel * (t - time_points[i - 1]))

    if t % t_dosage_interval == 0:
        # Konzentration durch Erhaltungsdosis nach einer Infusion von 1 Stunde
        inf_concentration = (dose_maintenance * (1 - np.exp(-kel * t_inf)) / (Vd * t_inf * kel))
        plasma_concentration[i] = previous_concentration + inf_concentration
    else:
        plasma_concentration[i] = previous_concentration

# Ergebnis plotten
plt.figure(figsize=(12, 6))
plt.plot(time_points, plasma_concentration, label="Vancomycin-Plasmaspiegel", color='b')
plt.xlabel("Zeit (Stunden)")
plt.ylabel("Plasmaspiegel (mg/L)")
plt.title("Pharmakokinetische Simulation der Vancomycin-Plasmaspiegel über 10 Tage")
plt.axhline(y=15, color='r', linestyle='--', label='Therapeutischer Bereich (15-20 mg/L)')
plt.axhline(y=20, color='r', linestyle='--')
plt.legend()
plt.grid()
plt.show()
