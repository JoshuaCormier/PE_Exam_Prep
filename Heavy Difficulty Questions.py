import csv
import random
import math

# --- Configuration ---
FILENAME = "fpe_pe_questions.csv"
NUM_QUESTIONS = 100

# --- Reference Data (Simulating Code Lookups) ---
# In a real exam, you look these up. Here, we simulate the "Step 1" lookup.
OCCUPANCY_LOAD_FACTORS = {
    "Business": 150,  # gross
    "Assembly (Unconcentrated)": 15,  # net
    "Assembly (Concentrated)": 7,  # net
    "Educational (Classroom)": 20,  # net
    "Mercantile (Basement)": 30,  # gross
}

PIPE_SCHEDULES = {
    "Schedule 40": {"id_inches": 4.026, "name": "4-inch Sched 40"},
    "Schedule 10": {"id_inches": 4.26, "name": "4-inch Sched 10"},
    "Schedule 40 (2in)": {"id_inches": 2.067, "name": "2-inch Sched 40"},
}


# --- Helper Functions ---
def generate_distractors(correct_value, variance=0.15):
    """Generates 3 plausible distractors based on percentage errors."""
    distractors = set()
    while len(distractors) < 3:
        # Create errors like "forgot square root" or "inverted fraction" or just random noise
        factor = random.uniform(1.0 - variance, 1.0 + variance)
        if factor == 1.0: continue

        val = correct_value * factor
        # Format similar to correct answer
        if isinstance(correct_value, int):
            val = int(val)
        else:
            val = round(val, 2)

        if val != correct_value and val > 0:
            distractors.add(val)
    return list(distractors)


# --- Question Archetypes ---

def q_hydraulics_friction_loss():
    # Hazen Williams: p = (4.52 * Q^1.85) / (C^1.85 * d^4.87) * L
    pipe = random.choice(list(PIPE_SCHEDULES.values()))
    c_factor = random.choice([100, 120, 140, 150])
    flow = random.randrange(250, 1000, 50)  # GPM
    length = random.randrange(10, 200, 10)  # Feet

    # Calculate Correct Answer
    d = pipe["id_inches"]
    friction_loss_psi = (4.52 * (flow ** 1.85)) / ((c_factor ** 1.85) * (d ** 4.87)) * length
    answer = round(friction_loss_psi, 2)

    # Create Question Text
    q_text = (f"Calculate the friction loss (psi) in a {length} ft section of {pipe['name']} steel pipe "
              f"with a C-factor of {c_factor} and a flow rate of {flow} gpm.")

    return "Hydraulics", q_text, answer, "psi"


def q_egress_capacity():
    # Step 1: Get Area & Occ Type -> Step 2: Get Load Factor -> Step 3: Calc Persons -> Step 4: Calc Width
    occ_type, factor = random.choice(list(OCCUPANCY_LOAD_FACTORS.items()))
    area = random.randrange(1000, 10000, 500)

    # Calculate Occupant Load
    occupants = math.ceil(area / factor)

    # Calculate Door Width (0.2 inches per person per IBC/NFPA 101 basic)
    required_width_inches = occupants * 0.2

    # Distractor Logic: Use 0.3 factor (stairs) instead of 0.2 (doors)
    wrong_stairs = round(occupants * 0.3, 2)

    answer = round(required_width_inches, 2)

    q_text = (f"A {area} sq.ft. {occ_type} space requires a main egress door. "
              f"Based on an occupant load factor of {factor} sq.ft./person and a capacity factor of 0.2 in/person, "
              f"what is the minimum total cumulative clear width (inches) required for the doors?")

    return "Egress", q_text, answer, "inches"


def q_smoke_axisymmetric():
    # Mass flow: m = 0.071 * Q_c^(1/3) * Z^(5/3) + ... (Simplified approx for Z > flame height)
    # Z = Ceiling Height - Layer Height (Interface)

    Q_c = random.randrange(1500, 5000, 500)  # kW (Convective)
    ceiling_h = random.choice([20, 30, 40])
    layer_h = random.choice([6, 10])  # Keep smoke above head

    z = ceiling_h - layer_h

    # Formula: m = 0.071 * (Q_c**(1/3)) * (z**(5/3))
    # (Note: Valid when z > mean flame height, assuming condition meets)
    m = 0.071 * (Q_c ** (1 / 3)) * (z ** (5 / 3))
    answer = round(m, 1)

    q_text = (f"Determine the mass flow rate of the smoke plume (kg/s) at a height of {layer_h} ft above the floor "
              f"in an atrium with a {ceiling_h} ft ceiling. The fire has a convective heat release rate of {Q_c} kW. "
              f"Assume the axisymmetric plume equation applies (m = 0.071 * Qc^(1/3) * z^(5/3)).")

    return "Smoke Control", q_text, answer, "kg/s"


def q_alarm_voltage_drop():
    # V_drop = (I * R * L) / 1000 (Simplified standard Ohm's law variants)
    # 14 AWG Solid Copper approx 2.57 ohm/1000ft (NEC Ch 9 Tbl 8)

    current = random.choice([0.5, 1.0, 2.0, 3.0])  # Amps
    length = random.randrange(100, 500, 50)  # Feet
    wire_r = 2.57  # Ohm/1000ft for #14

    # 2-wire circuit (Standard NAC) -> Length * 2
    v_drop = (current * wire_r * (length * 2)) / 1000
    answer = round(v_drop, 2)

    q_text = (f"A Notification Appliance Circuit (NAC) utilizes 14 AWG solid copper wire (R = 2.57 ohms/1000ft). "
              f"The circuit draws {current} Amps and the distance to the last appliance is {length} feet. "
              f"Calculate the voltage drop (Volts). Assume a standard 2-wire circuit.")

    return "Fire Alarm", q_text, answer, "Volts"


# --- Main Generation Loop ---

def generate_csv():
    archetypes = [q_hydraulics_friction_loss, q_egress_capacity, q_smoke_axisymmetric, q_alarm_voltage_drop]

    with open(FILENAME, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Header
        writer.writerow(["ID", "Domain", "Question", "Option A", "Option B", "Option C", "Option D", "Correct Answer",
                         "Explanation"])

        for i in range(1, NUM_QUESTIONS + 1):
            # Pick a random question type
            func = random.choice(archetypes)
            domain, text, correct, unit = func()

            # Generate options
            distractors = generate_distractors(correct)
            options = distractors + [correct]
            random.shuffle(options)

            # Map options to letters
            opt_map = {0: "A", 1: "B", 2: "C", 3: "D"}
            correct_letter = opt_map[options.index(correct)]

            # Form explanation
            explanation = f"Correct Answer: {correct} {unit}. Derived using standard engineering formulas for {domain}."

            writer.writerow([
                i,
                domain,
                text,
                f"{options[0]} {unit}",
                f"{options[1]} {unit}",
                f"{options[2]} {unit}",
                f"{options[3]} {unit}",
                correct_letter,
                explanation
            ])

    print(f"Successfully generated {NUM_QUESTIONS} questions in '{FILENAME}'")


if __name__ == "__main__":
    generate_csv()