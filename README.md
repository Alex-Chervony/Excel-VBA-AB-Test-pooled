VBA code and formulae allowing statistical analysis in MS Excel without add-ons.

# AB-Test
1. If test conditions\* are met, do a pooled Z test comparing 2 proportions.
1. If test conditions are met, return the proportion difference.
1. If test conditions are not met do not return anything.


### Usage:
```vba
=ABTest_Proportions(n1,y1,n2,y2)
```

### \* Test Conditions:
1. n1*p1 > 5 
1. n1*(1 - p1) > 5 
1. n2 * p2 > 5 
1. n2(1 - p2) > 5

Where:
1. n1 - sample size of group 1
1. n2 - sample size of group 2
1. y1 - count of instances of success in group 1
1. y2 - count of instances of success in group 2
1. p1 - proportion of group 1: y1/n1
1. p2 - proportion of group 2: y2/n2

