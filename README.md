# Particle Swarm Optimisation
Contains all necessary resources required to execute PSO using a python script and to document and the readings in a .xml database.

## Introduction
Particle swarm optimization (PSO) is one of the bio-inspired algorithm that searches for an optimal solution in the solution space. It is different from other optimization algorithms in the sense that only the objective function is needed and it is not dependent on the gradient or any differential form of the objective

## Description

Particle Swarm Optimisation begins with a initializing population (Similar to generic algorithms).
However unlike Generic algorithm, each particle is given a randomized velocity to explore the search space of its own accord.
_NOTE: Here a particle refers to a solution within the search space of the PSO._

### The 3 distinct features of PSO are
1. **Best fitness of a particle**: The best solution achieved so far by a particular particle _i_ (i.e Local Best).
2. **Best fitness of the swarm**: The best solution achieved so far by any particel in the swarm (i.e Global Best).
3. **Velocity and position update of each particle**: For exploring and exploiting the search space to loacte the optimal solution.

## Process Flow

![Screenshot 2023-07-24 234105](https://github.com/NakulSK221B/Particle-Swarm-Optimisation/assets/95758559/be035f8b-403f-42b9-b0ad-5bcf04d07c52)

## Results
The result of the process is first displayed in the form of a plot using the _matplotlib_ library.

![image](https://github.com/NakulSK221B/Particle-Swarm-Optimisation/assets/95758559/51629105-eae8-406c-a8dc-774106cb601c)

In the above image the _triangles_ represent the initializing population of particles.
The _squares_ that gradually accumalate towards each other are the generation of particles that eventually coincide towards the optimal solution.

After this, The detailed information of all the iteration of the algorithm are documented in the Test_WB.xlsx file.

![image](https://github.com/NakulSK221B/Particle-Swarm-Optimisation/assets/95758559/018c1ef7-c80c-4537-b00c-2411de303b19)

https://github.com/NakulSK221B/Particle-Swarm-Optimisation/blob/4f8d00ef448308a474e1954ecb2d9498ad0a4a22/v3.0/Test_WB.xlsx



