# N-Queen Solver

## Description

The N-Queen Solver is a Python script that solves the N-Queen problem using backtracking. The N-Queen problem is the challenge of placing N chess queens on an N×N chessboard so that no two queens threaten each other; thus, a solution requires that no two queens share the same row, column, or diagonal.

## Dependencies

The script requires the following Python libraries:
- `openpyxl`: This library is used to manipulate Excel files. You can install it via pip:
  ```
  pip install openpyxl
  ```

## Usage

1. Clone the repository or download the `N_Queen.py` file.
2. Ensure you have Python installed on your system.
3. Run the script in your preferred Python environment.

## Note

The script is designed to handle N-Queen puzzles where N is between 4 and 15. While it can technically handle larger values of N, the performance may degrade significantly, and the solution may take longer to compute due to backtracking algorithm.
