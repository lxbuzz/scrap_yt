import csv

queries_data = [
    {"ID": "A2.7", "Owner": "P2", "Theme": "A2 Low appetite and under-eating", "Query type": "Pain", "Query (copy as written)": "what I eat in a day on Zepbound", "Platform": "YouTube", "Priority": 1},
    {"ID": "A1.7", "Owner": "P2", "Theme": "A1 Protein and muscle preservation", "Query type": "Pain", "Query (copy as written)": "high protein meal prep for GLP-1", "Platform": "YouTube", "Priority": 2},
    {"ID": "B1.7", "Owner": "P2", "Theme": "B1 Dinner decision fatigue", "Query type": "Pain", "Query (copy as written)": "what's for dinner this week family of 4", "Platform": "YouTube", "Priority": 2},
    {"ID": "B2.7", "Owner": "P2", "Theme": "B2 Picky eaters and kids who refuse", "Query type": "Pain", "Query (copy as written)": "picky eater family meals", "Platform": "YouTube", "Priority": 2},
    {"ID": "A1.8", "Owner": "P2", "Theme": "A1 Protein and muscle preservation", "Query type": "Pain", "Query (copy as written)": "how to get 100g protein on Zepbound", "Platform": "YouTube", "Priority": 3},
    {"ID": "A4.7", "Owner": "P2", "Theme": "A4 Small portions, waste, cooking for one", "Query type": "Pain", "Query (copy as written)": "GLP-1 meal prep small portions", "Platform": "YouTube", "Priority": 3},
    {"ID": "A5.6", "Owner": "P2", "Theme": "A5 Cooking for a household while on GLP-1", "Query type": "Pain", "Query (copy as written)": "family dinner GLP-1 meal ideas", "Platform": "YouTube", "Priority": 3},
]

with open("queries_config.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=["ID", "Owner", "Theme", "Query type", "Query (copy as written)", "Platform", "Priority"])
    writer.writeheader()
    writer.writerows(queries_data)

print("Saved queries_config.csv")
