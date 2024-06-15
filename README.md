
# C++ Spreadsheet Class

## Overview
This project involves creating a class (or a set of classes) that functions as a spreadsheet processor. The spreadsheet processor allows operations on cells (setting, computing values, copying), calculates the value of a cell based on a formula, detects cyclic dependencies between cells, and can save and load the spreadsheet's content. The design focuses on appropriate class design, the use of polymorphism, and effective use of a version control system.

## Features
- Cell operations including setting values, calculating based on formulas, and copying.
- Detection of cyclic dependencies to prevent infinite loops.
- Ability to save and load the spreadsheet state.
- Integration with a provided expression parser in the form of a statically linked library.

## Technologies
- C++
- Statically linked libraries for expression parsing.

## Getting Started

### Prerequisites
- C++17 or higher.
- Development environment capable of C++ compilation.
- Access to a version control system (e.g., Git).

### Building the Project
To build the project, run the following command in the root directory of the project:
```bash
g++ -o Spreadsheet main.cpp
```
Make sure to replace `main.cpp` with the actual file names of your source code.

### Usage
After building the project, you can run the executable:
```bash
./Spreadsheet
```

## Contact
For any further queries, you can reach out at matus.gls.tkac@gmail.com.
