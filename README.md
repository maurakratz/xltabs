
# xltabs

The goal of xltabs is to generate professional cross-tabulations and
frequency tables and export them to Excel with proper formatting. The
output mimics the visual layout of Stata tables by combining counts and
percentages into a single cell using line breaks. The display of counts,
row and column totals, missing values, weights, and percentages is
customisable. Tables can be generated from raw or from pre-aggregated
data (e.g., in case of complex survey design).

## Installation

You can install the development version of xltabs from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("maurakratz/xltabs")
```

## Example

Here is a basic example of how you can create a Stata-styled crosstable
for export to Excel with xltabs:

``` r
library(xltabs)
library(dplyr)

# 1. Create dummy data
df <- tibble(
  edu = c("High", "Low", "High", "Low", "High"),
  job = c("Yes", "No", "Yes", "Yes", "No"),
  weight = c(1, 1.2, 0.8, 1, 1)
)

# 2. Create the crosstab data frame
my_tab <- xl_crosstab(df, edu, job, w_var = weight, 
                      row_label = "Education", col_label = "Employed")

# 3. Print it (Stata style!)
print(my_tab)
#> # A tibble: 3 × 4
#>   Education No                       Yes                      Total     
#>   <fct>     <chr>                    <chr>                    <chr>     
#> 1 High      "1\n35.7%\n45.5%\n20.0%" "2\n64.3%\n64.3%\n36.0%" "3\n56.0%"
#> 2 Low       "1\n54.5%\n54.5%\n24.0%" "1\n45.5%\n35.7%\n20.0%" "2\n44.0%"
#> 3 Total     "2\n44.0%"               "3\n56.0%"               "5"
```

To export this table to an Excel sheet with title and footer:

``` r
# 1 Export to Excel
xl_export(my_tab, "report.xlsx", 
          title = "Table 1: Education by Employment",
          footer = "Source: Toy Data")

# 2 Open the report.xlsx file to check out the layout (Don't forget to enable text wrapping in Excel for the line breaks!)

# 3 You might want to delete the file created by this example so it doesn't clutter your folder
if (file.exists("report.xlsx")) file.remove("report.xlsx")
#> [1] TRUE
```
