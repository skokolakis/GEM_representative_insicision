# GEM_representative_insicision
This Python script processes GEM-2 SKI survey data collected along stream transects using multiple frequencies. 
It is designed to work with Excel files where each sheet corresponds to a frequency, and each column represents a separate measurement line along the stream.

This script processes multiple Excel files containing longitudinal electrical conductivity (EC) profiles, interpolates the data to a common distance axis, computes representative profiles for each sheet (frequency), and ranks them based on a representativeness score.
Workflow:
1. Reads all `.xlsx` files in the specified input folder, skipping temporary files.
2. For each Excel file:
    - Iterates through each sheet (assumed to represent different frequencies or measurement lines).
    - Extracts the distance column and EC response columns.
    - Interpolates all numeric EC columns to a common distance axis using the specified interpolation method.
    - Computes the mean (representative profile) and standard deviation across all interpolated lines.
    - Calculates a representativeness score for each sheet: (amplitude of mean profile) / (mean standard deviation).
    - Stores interpolated mean profiles for output.
    - Plots all representative profiles for visual comparison, labeling with their scores.
3. Saves the interpolated mean profiles to a new Excel file in the output folder.
4. Saves the plot of representative profiles as a PNG image in the output folder.
5. Prints a ranking of all sheets (frequencies) in each file by their representativeness score, highlighting the best one.
Configuration:
- `common_distance_step`: Step size for the common distance axis (meters).
- `interp_kind`: Interpolation method (e.g., 'linear', 'nearest', etc.).
- `input_folder`: Folder containing input Excel files.
- `output_folder`: Folder to save output Excel files and plots.
- `y_axis_label`: Label for the y-axis in plots.
Assumptions:
- The first column in each sheet is the distance (meters).
- Remaining columns are numeric EC responses.
- Sheets with insufficient or invalid data are skipped.
Dependencies:
- pandas, numpy, matplotlib, scipy, os
Outputs:
- Interpolated mean profiles per sheet as Excel files.
- Plots of representative profiles.
- Console ranking of frequencies by representativeness score.
