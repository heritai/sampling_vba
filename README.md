# Sampling Methodes in VBA(Excel)

This repository contains VBA functions for performing statistical sampling analysis in Microsoft Excel. These functions are designed to automate the calculation of key statistical parameters, such as the mean and variance, for various sampling methods.

## Purpose

The goal of this project is to provide reusable VBA functions that simplify and streamline the process of statistical analysis within Excel. By using these functions, users can avoid manual calculations and quickly obtain accurate results for different sampling techniques.

## VBA Functions

This repository includes the following VBA functions:

*   **`cls(R1 As Range, R2 As Range, what As Byte) As Double`**

    *   Calculates the mean or variance for stratified sampling.

    *   **Parameters:**

        *   `R1`: A `Range` object containing the sample data. It should have two columns:
         the identifier of the sample and the value of the metric.
        *   `R2`: A `Range` object containing the strata data. It should have three columns:
         the identifier of the strata, total number of the samples and the value of the strata.
        *   `what`: A `Byte` value indicating whether to calculate the mean (1) or variance (anything else).

    *   **Return Value:**

        *   The calculated mean or variance as a `Double`.
*   **`clts(R1 As Range, R2 As Range, N As Long, m As Long, what As Byte) As Double`**

    *   Calculates the mean or variance for cluster sampling.

    *   **Parameters:**

        *   `R1`: A `Range` object containing the sample data. It should have two columns:
         the identifier of the sample and the value of the metric.
        *   `R2`: A `Range` object containing the cluster data. It should have two columns:
         the identifier of the cluster and the total size of the cluster.
        *   `N`: A `Long` value representing the population size.
        *   `m`: A `Long` value representing the number of clusters.
        *   `what`: A `Byte` value indicating whether to calculate the mean (1) or variance (anything else).

    *   **Return Value:**

        *   The calculated mean or variance as a `Double`.

*   **`cltsInClss(R1 As Range, R2 As Range, R3 As Range, what As Byte) As Double`**

*  Calculates the mean or variance for cluster inside class sampling

*   **Parameters:**

     *   `R1`: A `Range` object containing the sample data. It should have three columns in order: the identifier of the sample, the id of the cluster and the value of the metric.
    *   `R2`: A `Range` object containing the data for the clusters. It should have three columns in order: the identifier of the cluster,  the value and the size of the cluster.
     *   `R3`: A `Range` object containing the data for the classes. It should have three columns in order: the identifier of the class, a value and the size of the class.
    *   `what`: A `Byte` value indicating whether to calculate the mean (1) or variance (anything else).

*   **Return Value:**

    *   The calculated mean or variance as a `Double`.

*   **`clssInClts(R1 As Range, R2 As Range, R3 As Range, N As Double, M As Double, what As Byte) As Double`**

    *   Calculates the mean or variance for stratified sampling inside a cluster.

    *   **Parameters:**

        *   `R1`: A `Range` object containing the sample data. It should have three columns in order: the identifier of the sample, the id of the cluster and the value of the metric.
        *   `R2`: A `Range` object containing the data for the clusters. It should have three columns in order: the identifier of the cluster, a value and the size of the cluster.
        *   `R3`: A `Range` object containing the data for the classes. It should have three columns in order: the identifier of the class, a value and the size of the class.
        *   `N`: The size of the population.
        *   `M`: The total number of the classes.
        *   `what`: A `Byte` value indicating whether to calculate the mean (1) or variance (anything else).

    *   **Return Value:**

        *   The calculated mean or variance as a `Double`.

## Usage Instructions

1.  **Open Microsoft Excel.**

2.  **Open the VBA Editor:** Press `Alt + F11`.

3.  **Insert a New Module:** In the VBA Editor, go to `Insert` \> `Module`.

4.  **Copy and Paste the Code:** Copy the contents of the desired `.vba` file (e.g., `classified.vba`) into the newly created module.

5.  **Close the VBA Editor:** Close the VBA Editor window.

6.  **Use the Function in Your Worksheet:**

    *   In your Excel worksheet, you can now use the VBA function like any other built-in Excel function.
    *   For example, if you pasted the code from `classified.vba`, and your data is arranged in the format for that function, you can use it like this in a cell:

        `=cls(A1:B100,D1:F5,1)`
        (Replace `A1:B100` and `D1:F5` with your actual data ranges).

## Data Organization

It's crucial to organize your data in the Excel worksheet according to the expected format for each function. Review the code comments and parameter descriptions carefully to understand the required data layout.

## Disclaimer

These VBA functions are provided as-is and may require adjustments to suit your specific data and analysis needs. Always verify the results to ensure accuracy.

## Author

Yousef Taheri.
