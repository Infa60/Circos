# bornes
old_min, old_max = 1, 53
new_min, new_max = 70, 400

# fonction de rééchelonnage
def rescale(x):
    return new_min + (x - old_min) * (new_max - new_min) / (old_max - old_min)

# exemple d'utilisation
list_to_rescale = list(range(1, 54))
scaled_values = [int(rescale(v)) for v in list_to_rescale]

for original, scaled in zip(list_to_rescale, scaled_values):
    print(original, scaled)