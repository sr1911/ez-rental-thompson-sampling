"""
EZ Car Rental: Dynamic Pricing with Thompson Sampling
ISBA 2415 - Reinforcement Learning - M5 Project Paper

Implements Thompson Sampling and epsilon-greedy agents for the EZ Car Rental
dynamic pricing problem, then compares their performance against an oracle.

The Journeys and Utilization CSVs serve as ground truth to build a simulator.
The agents have NO direct access to this data.
"""

import csv
import random
import os
from collections import defaultdict
from datetime import datetime

random.seed(42)

# =============================================================================
# 1. LOAD AND PROCESS DATA
# =============================================================================

BASE = os.path.dirname(os.path.abspath(__file__))
JOURNEYS_PATH = os.path.join(BASE, "journeys.csv")

def load_journeys(path):
    rows = []
    with open(path, "r") as f:
        reader = csv.DictReader(f)
        for row in reader:
            price = float(row["Trip Sum Trip Price"].replace("$", "").replace(",", ""))
            start = datetime.strptime(row["Trip Start At Local Time"], "%Y-%m-%d %H:%M:%S")
            end = datetime.strptime(row["Trip End At Local Time"], "%Y-%m-%d %H:%M:%S")
            duration_hrs = max((end - start).total_seconds() / 3600, 0.5)
            hourly_rate = price / duration_hrs
            city = row["Car Parking Address City"].strip()
            hour = start.hour
            rows.append({
                "city": city,
                "hour": hour,
                "hourly_rate": hourly_rate,
            })
    return rows

print("Loading journeys data...")
journeys = load_journeys(JOURNEYS_PATH)
print(f"  Loaded {len(journeys)} trips")

city_counts = defaultdict(int)
for j in journeys:
    city_counts[j["city"]] += 1
CITIES = sorted(city_counts, key=city_counts.get, reverse=True)[:5]
print(f"  Top 5 cities: {CITIES}")

journeys = [j for j in journeys if j["city"] in CITIES]
print(f"  Filtered to {len(journeys)} trips in top 5 cities")

# =============================================================================
# 2. STATE AND ACTION SPACES
# =============================================================================

PRICES = [3, 5, 7, 9, 11]
NUM_ACTIONS = len(PRICES)
city_to_idx = {c: i for i, c in enumerate(CITIES)}
NUM_STATES = len(CITIES) * 2 * 2  # 20

def state_tuple(city_idx, time_cat, util_level):
    return (city_idx, time_cat, util_level)

def state_to_idx(s):
    return s[0] * 4 + s[1] * 2 + s[2]

def idx_to_state(idx):
    return (idx // 4, (idx % 4) // 2, idx % 2)

def state_label(s):
    return f"({CITIES[s[0]]}, {'Peak' if s[1] else 'Off-peak'}, {'High-util' if s[2] else 'Low-util'})"

# =============================================================================
# 3. BUILD SIMULATOR
# =============================================================================

def compute_empirical_booking_rates(journeys):
    groups = defaultdict(list)
    for j in journeys:
        city_idx = city_to_idx[j["city"]]
        time_cat = 1 if 7 <= j["hour"] <= 21 else 0
        groups[(city_idx, time_cat)].append(j["hourly_rate"])

    booking_probs = {}
    for (city_idx, time_cat), rates in groups.items():
        rates_sorted = sorted(rates)
        median_rate = rates_sorted[len(rates_sorted) // 2]
        low_util_rates = [r for r in rates if r < median_rate]
        high_util_rates = [r for r in rates if r >= median_rate]

        for util_level, subset in [(0, low_util_rates), (1, high_util_rates)]:
            if not subset:
                subset = rates
            for price in PRICES:
                n_willing = sum(1 for r in subset if r >= price)
                prob = max(0.05, min(0.95, n_willing / len(subset)))
                s = state_tuple(city_idx, time_cat, util_level)
                booking_probs[(s, price)] = prob
    return booking_probs

print("\nBuilding simulator...")
TRUE_BOOKING_PROBS = compute_empirical_booking_rates(journeys)

def simulate_booking(state, price, rng=None):
    prob = TRUE_BOOKING_PROBS.get((state, price), 0.3)
    r = (rng or random).random()
    return 1 if r < prob else 0

# =============================================================================
# 4. AGENTS
# =============================================================================

class ThompsonSamplingAgent:
    def __init__(self):
        self.alpha = {}
        self.beta_param = {}
        for si in range(NUM_STATES):
            s = idx_to_state(si)
            for a in PRICES:
                self.alpha[(s, a)] = 1.0
                self.beta_param[(s, a)] = 1.0

    def select_action(self, state, rng=None):
        r = rng or random
        best_action, best_value = None, -1
        for a in PRICES:
            theta = r.betavariate(self.alpha[(state, a)], self.beta_param[(state, a)])
            ev = a * theta
            if ev > best_value:
                best_value = ev
                best_action = a
        return best_action

    def update(self, state, action, reward):
        self.alpha[(state, action)] += reward
        self.beta_param[(state, action)] += (1 - reward)

    def get_learned_probs(self):
        return {k: self.alpha[k] / (self.alpha[k] + self.beta_param[k]) for k in self.alpha}

    def get_policy(self):
        policy = {}
        for si in range(NUM_STATES):
            s = idx_to_state(si)
            best_price, best_rev = None, -1
            for a in PRICES:
                mean_theta = self.alpha[(s, a)] / (self.alpha[(s, a)] + self.beta_param[(s, a)])
                ev = a * mean_theta
                if ev > best_rev:
                    best_rev = ev
                    best_price = a
            policy[s] = best_price
        return policy


class EpsilonGreedyAgent:
    def __init__(self, epsilon=0.1):
        self.epsilon = epsilon
        self.counts = {}
        self.sum_rewards = {}
        for si in range(NUM_STATES):
            s = idx_to_state(si)
            for a in PRICES:
                self.counts[(s, a)] = 0
                self.sum_rewards[(s, a)] = 0.0

    def select_action(self, state, rng=None):
        r = rng or random
        if r.random() < self.epsilon:
            return r.choice(PRICES)
        best_action, best_rev = None, -1
        for a in PRICES:
            n = self.counts[(state, a)]
            if n == 0:
                return a  # try untried actions first
            mean_reward = self.sum_rewards[(state, a)] / n
            ev = a * mean_reward
            if ev > best_rev:
                best_rev = ev
                best_action = a
        return best_action

    def update(self, state, action, reward):
        self.counts[(state, action)] += 1
        self.sum_rewards[(state, action)] += reward

    def get_policy(self):
        policy = {}
        for si in range(NUM_STATES):
            s = idx_to_state(si)
            best_price, best_rev = None, -1
            for a in PRICES:
                n = self.counts[(s, a)]
                if n == 0:
                    continue
                mean_reward = self.sum_rewards[(s, a)] / n
                ev = a * mean_reward
                if ev > best_rev:
                    best_rev = ev
                    best_price = a
            policy[s] = best_price if best_price else PRICES[0]
        return policy


# =============================================================================
# 5. ORACLE
# =============================================================================

def compute_oracle_policy():
    policy, oracle_rev = {}, {}
    for si in range(NUM_STATES):
        s = idx_to_state(si)
        best_price, best_rev = None, -1
        for p in PRICES:
            rev = p * TRUE_BOOKING_PROBS.get((s, p), 0)
            if rev > best_rev:
                best_rev = rev
                best_price = p
        policy[s] = best_price
        oracle_rev[s] = best_rev
    return policy, oracle_rev

ORACLE_POLICY, ORACLE_REV = compute_oracle_policy()


# =============================================================================
# 6. RUN BOTH AGENTS WITH SAME RANDOM SEQUENCE
# =============================================================================

NUM_EPISODES = 5000
EVAL_INTERVAL = 100

def run_agent(agent_class, seed=42, **kwargs):
    rng = random.Random(seed)
    agent = agent_class(**kwargs) if kwargs else agent_class()

    episode_revenues = []
    episode_oracle_revenues = []
    running_rev = 0
    running_oracle = 0
    regret = []
    policy_accuracy = []

    for ep in range(NUM_EPISODES):
        si = rng.randint(0, NUM_STATES - 1)
        state = idx_to_state(si)

        price = agent.select_action(state, rng=rng)
        booked = simulate_booking(state, price, rng=rng)

        rev = price * booked
        oracle_rev_ep = ORACLE_REV[state]

        agent.update(state, price, booked)

        running_rev += rev
        running_oracle += oracle_rev_ep
        episode_revenues.append(rev)
        episode_oracle_revenues.append(oracle_rev_ep)
        regret.append(running_oracle - running_rev)

        if (ep + 1) % EVAL_INTERVAL == 0:
            learned = agent.get_policy()
            matches = sum(1 for si2 in range(NUM_STATES)
                         if learned[idx_to_state(si2)] == ORACLE_POLICY[idx_to_state(si2)])
            policy_accuracy.append((ep + 1, matches / NUM_STATES))

    return {
        "agent": agent,
        "episode_revenues": episode_revenues,
        "episode_oracle_revenues": episode_oracle_revenues,
        "regret": regret,
        "policy_accuracy": policy_accuracy,
        "total_revenue": sum(episode_revenues),
        "total_oracle": running_oracle,
    }

print("\nRunning Thompson Sampling...")
ts_results = run_agent(ThompsonSamplingAgent, seed=42)
print(f"  Revenue: ${ts_results['total_revenue']:,.2f} / ${ts_results['total_oracle']:,.2f} "
      f"({ts_results['total_revenue']/ts_results['total_oracle']:.1%})")

print("Running Epsilon-Greedy (eps=0.15)...")
eg_results = run_agent(EpsilonGreedyAgent, seed=42, epsilon=0.15)
print(f"  Revenue: ${eg_results['total_revenue']:,.2f} / ${eg_results['total_oracle']:,.2f} "
      f"({eg_results['total_revenue']/eg_results['total_oracle']:.1%})")


# =============================================================================
# 7. PRINT RESULTS
# =============================================================================

print("\n" + "=" * 80)
print("COMPARISON RESULTS")
print("=" * 80)

ts_policy = ts_results["agent"].get_policy()
eg_policy = eg_results["agent"].get_policy()

ts_match = sum(1 for si in range(NUM_STATES) if ts_policy[idx_to_state(si)] == ORACLE_POLICY[idx_to_state(si)])
eg_match = sum(1 for si in range(NUM_STATES) if eg_policy[idx_to_state(si)] == ORACLE_POLICY[idx_to_state(si)])

print(f"\n  {'Metric':<35} {'Thompson':>12} {'Eps-Greedy':>12}")
print("  " + "-" * 60)
print(f"  {'Total Revenue':<35} ${ts_results['total_revenue']:>10,.2f} ${eg_results['total_revenue']:>10,.2f}")
print(f"  {'Oracle Revenue':<35} ${ts_results['total_oracle']:>10,.2f} ${eg_results['total_oracle']:>10,.2f}")
print(f"  {'Revenue Ratio':<35} {ts_results['total_revenue']/ts_results['total_oracle']:>11.1%} {eg_results['total_revenue']/eg_results['total_oracle']:>11.1%}")
print(f"  {'Final Regret':<35} ${ts_results['regret'][-1]:>10,.2f} ${eg_results['regret'][-1]:>10,.2f}")
print(f"  {'Policy Accuracy':<35} {ts_match}/{NUM_STATES:>9} {eg_match}/{NUM_STATES:>9}")

# Learned probs error
ts_probs = ts_results["agent"].get_learned_probs()
ts_errors = [abs(ts_probs[(idx_to_state(si), p)] - TRUE_BOOKING_PROBS.get((idx_to_state(si), p), 0))
             for si in range(NUM_STATES) for p in PRICES]
print(f"  {'Mean Abs Error (all pairs)':<35} {sum(ts_errors)/len(ts_errors):>12.4f} {'N/A':>12}")

print("\n--- Final Policy Comparison ---")
print(f"  {'State':<40} {'TS':>5} {'EG':>5} {'Oracle':>7}")
print("  " + "-" * 60)
for si in range(NUM_STATES):
    s = idx_to_state(si)
    print(f"  {state_label(s):<40} ${ts_policy[s]:>3} ${eg_policy[s]:>3}  ${ORACLE_POLICY[s]:>3}")


# =============================================================================
# 8. EXPORT ALL RESULTS
# =============================================================================

# Regret comparison
with open(os.path.join(BASE, "results_regret.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Episode", "TS_Regret", "EG_Regret"])
    for i in range(0, NUM_EPISODES, 10):
        writer.writerow([i+1, round(ts_results["regret"][i], 2), round(eg_results["regret"][i], 2)])

# Policy accuracy comparison
with open(os.path.join(BASE, "results_policy_accuracy.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Episode", "TS_Accuracy", "EG_Accuracy"])
    for (ep1, acc1), (ep2, acc2) in zip(ts_results["policy_accuracy"], eg_results["policy_accuracy"]):
        writer.writerow([ep1, round(acc1, 4), round(acc2, 4)])

# Windowed average revenue
window = 200
with open(os.path.join(BASE, "results_avg_revenue.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Window_Start", "TS_Revenue", "EG_Revenue", "Oracle_Revenue"])
    for i in range(0, NUM_EPISODES - window + 1, window):
        ts_avg = sum(ts_results["episode_revenues"][i:i+window]) / window
        eg_avg = sum(eg_results["episode_revenues"][i:i+window]) / window
        or_avg = sum(ts_results["episode_oracle_revenues"][i:i+window]) / window
        writer.writerow([i+1, round(ts_avg, 4), round(eg_avg, 4), round(or_avg, 4)])

# Policy table
with open(os.path.join(BASE, "results_policy.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["State", "TS_Price", "EG_Price", "Oracle_Price", "TS_Match", "EG_Match"])
    for si in range(NUM_STATES):
        s = idx_to_state(si)
        writer.writerow([
            state_label(s), ts_policy[s], eg_policy[s], ORACLE_POLICY[s],
            "Y" if ts_policy[s] == ORACLE_POLICY[s] else "N",
            "Y" if eg_policy[s] == ORACLE_POLICY[s] else "N",
        ])

# Probabilities
with open(os.path.join(BASE, "results_probabilities.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["State", "Price", "Learned_Prob", "True_Prob", "Abs_Error"])
    for si in range(NUM_STATES):
        s = idx_to_state(si)
        for p in PRICES:
            learned = ts_probs[(s, p)]
            true_p = TRUE_BOOKING_PROBS.get((s, p), 0)
            writer.writerow([state_label(s), p, round(learned, 4), round(true_p, 4), round(abs(learned - true_p), 4)])

print("\nResults exported to CSV files.")
print("Done!")
