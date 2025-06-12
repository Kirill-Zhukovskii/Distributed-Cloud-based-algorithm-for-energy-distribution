import pandas as pd
import matplotlib.pyplot as plt
import random
from datetime import datetime, time

# Helper function
def extract_hour(value):
    try:
        if isinstance(value, str):
            value = pd.to_datetime(value)
        if isinstance(value, (datetime, pd.Timestamp)):
            return value.hour + value.minute / 60
        elif isinstance(value, time):
            return value.hour + value.minute / 60
        elif isinstance(value, (int, float)):
            return float(value)
    except Exception:
        pass
    raise ValueError(f"Unsupported time format: {value}")

# EV class
class ElectricVehicle:
    def __init__(self, profile_row, battery_capacity=10000):
        self.arrival = extract_hour(profile_row['arrivalAtHome'])
        self.departure = extract_hour(profile_row['departureFromHome'])
        self.consumption = float(profile_row['consumption'])
        self.capacity = battery_capacity
        self.initial_soc = random.uniform(0.1, 0.2) * self.capacity
        self.soc = self.initial_soc

    def reset_day(self):
        """Reset SoC randomly for new day."""
        self.initial_soc = random.uniform(0.1, 0.2) * self.capacity
        self.soc = self.initial_soc

    def simulate_day(self, time_step=0.25):
        """Simulate one day with arrival, departure, and charging."""
        self.soc -= self.consumption
        self.soc = max(self.soc, 0)

        charging_duration = (self.departure - self.arrival) % 24
        steps = int(charging_duration / time_step)

        goal_soc = self.consumption
        power_per_step = (goal_soc - self.soc) / charging_duration if charging_duration > 0 else 0

        for _ in range(steps):
            if self.soc >= self.capacity:
                break
            self.soc += min(power_per_step * time_step, self.capacity - self.soc)

        return {
            'Initial SoC': round(self.initial_soc, 2),
            'Goal (kWh)': round(self.consumption, 2),
            'Arrival (h)': round(self.arrival, 2),
            'Departure (h)': round(self.departure, 2)
        }

# Fleet manager
class EVFleet:
    def __init__(self, xlsx_path, num_vehicles=10):
        df = pd.read_excel(xlsx_path)
        self.profiles = df
        self.vehicles = [ElectricVehicle(df.sample(n=1).iloc[0]) for _ in range(num_vehicles)]

    def simulate_multiple_days(self, num_days=1):
        all_results = []
        for day in range(num_days):
            for i, ev in enumerate(self.vehicles):
                ev.reset_day()
                daily_result = ev.simulate_day()
                daily_result.update({'EV': i, 'Day': day + 1})
                all_results.append(daily_result)
        return all_results