# Excel-Statistical-Functions-Add-in
XLS Statistics Functions Excel Add-In
<!-- wp:paragraph -->
<p>This Excel addin adds the following new functions to excel:</p>
<!-- /wp:paragraph -->

<!-- wp:list -->
<ul><li>DOWNSIDEDEVIATION</li><li>GAINLOSSRATIO</li><li>GAINSTANDARDDEVIATION</li><li>LOSSSTANDARDDEVIATION</li><li>SEMIDEVIATION</li><li>COMPOUNDRETURN</li><li>GAINMEAN</li><li>LOSSMEAN</li><li>COVARIANCE</li><li>FIVENUMBERSUMMARYINPLACE</li><li>GEOMETRICMEAN</li><li>HARMONICMEAN</li><li>INTERQUARTILERANGEINPLACE</li><li>LOWERQUARTILEINPLACE</li><li>MAXIMUM</li><li>MAXIMUMABSOLUTE</li><li>MEAN</li><li>MEANSTANDARDDEVIATION</li><li>MEANVARIANCE</li><li>MEDIANINPLACE</li><li>MINIMUM</li><li>MINIMUMABSOLUTE</li><li>ORDERSTATISTICINPLACE</li><li>PERCENTILEINPLACE</li><li>POPULATIONCOVARIANCE</li><li>POPULATIONSTANDARDDEVIATION</li><li>POPULATIONVARIANCE</li><li>QUANTILECUSTOMINPLACE</li><li>QUANTILEINPLACE</li><li>RANKSINPLACE</li><li>ROOTMEANSQUARE</li><li>STANDARDDEVIATION</li><li>UPPERQUARTILEINPLACE</li><li>VARIANCE</li></ul>
<!-- /wp:list -->

<!-- wp:paragraph -->
<p><a href="https://thedigitalbrain.ch/index.php/sdm_downloads/xls-statistics-functions-excel-add-in/" target="_blank" rel="noreferrer noopener">Click Here to install</a></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Below you will find a description of all additional functions:</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>DOWNSIDEDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>This measure is similar to the loss standard deviation except the downside deviation considers only returns that fall below a defined minimum acceptable return (MAR) rather than the arithmetic mean. For example, if the MAR is 7%, the downside deviation would measure the variation of each period that falls below 7%. (The loss standard deviation, on the other hand, would take only losing periods, calculate an average return for the losing periods, and then measure the variation between each losing return and the losing return average).</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>GAINLOSSRATIO</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Measures a fund’s average gain in a gain period divided by the fund’s average loss in a losing period. Periods can be monthly or quarterly depending on the data frequency.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>GAINSTANDARDDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Calculation is similar to Standard Deviation , except it calculates an average (mean) return only for periods with a gain and measures the variation of only the gain periods around the gain mean. Measures the volatility of upside performance. © Copyright 1996, 1999 Gary L.Gastineau. First Edition. © 1992 Swiss Bank Corporation.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>LOSSSTANDARDDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Similar to standard deviation, except this statistic calculates an average (mean) return for only the periods with a loss and then measures the variation of only the losing periods around this loss mean. This statistic measures the volatility of downside performance.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>SEMIDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>A measure of volatility in returns below the mean. It's similar to standard deviation, but it only looks at periods where the investment return was less than average return.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>COMPOUNDRETURN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Compound Monthly Return or Geometric Return or Annualized Return</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>GAINMEAN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Average Gain or Gain Mean This is a simple average (arithmetic mean) of the periods with a gain. It is calculated by summing the returns for gain periods (return 0) and then dividing the total by the number of gain periods.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>LOSSMEAN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Average Loss or LossMean This is a simple average (arithmetic mean) of the periods with a loss. It is calculated by summing the returns for loss periods (return &lt; 0) and then dividing the total by the number of loss periods.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>COVARIANCE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the unbiased population covariance from the provided two sample arrays. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>FIVENUMBERSUMMARYINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates {min, lower-quantile, median, upper-quantile, max} from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>GEOMETRICMEAN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the geometric mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>HARMONICMEAN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the harmonic mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>INTERQUARTILERANGEINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the inter-quartile range from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>LOWERQUARTILEINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the first quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MAXIMUM</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Returns the largest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MAXIMUMABSOLUTE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Returns the largest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MEAN</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the arithmetic sample mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MEANSTANDARDDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the arithmetic sample mean and the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for standard deviation if data has less than two entries or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MEANVARIANCE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the arithmetic sample mean and the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for variance if data has less than two entries or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MEDIANINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the median value from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MINIMUM</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Returns the smallest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>MINIMUMABSOLUTE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Returns the smallest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>ORDERSTATISTICINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Returns the order statistic (order 1..N) from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>PERCENTILEINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the p-Percentile value from the unsorted data array. If a non-integer Percentile is needed, use Quantile instead. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>POPULATIONCOVARIANCE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the population covariance from the full population provided as two arrays. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>POPULATIONSTANDARDDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the population standard deviation from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>POPULATIONVARIANCE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the population variance from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>QUANTILECUSTOMINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. The quantile definition can be specified by 4 parameters a, b, c and d, consistent with Mathematica. WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>QUANTILEINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>RANKSINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Evaluates the rank of each entry of the unsorted data array. The rank definition can be specified to be compatible with an existing system. WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>ROOTMEANSQUARE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the root mean square (RMS) also known as quadratic mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>STANDARDDEVIATION</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>UPPERQUARTILEINPLACE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the third quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.</p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p><strong>VARIANCE</strong></p>
<!-- /wp:paragraph -->

<!-- wp:paragraph -->
<p>Estimates the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.</p>
<!-- /wp:paragraph -->
