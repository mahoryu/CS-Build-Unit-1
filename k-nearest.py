#k-nearest.py
from math import sqrt

# 3 steps:

# 1. Calculate Euclidean Distance
def euclidean_distance(row1,row2):
    dist = 0.0
    for i in range(len(row1)-1):
        dist += (row1[i] - row2[i])**2
    return sqrt(dist)

# 2. Get Nearest Neighbors
# 3. Make Predictions



# print(euclidean_distance([1,2,3],[3,4,5]))