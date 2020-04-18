using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace XLSStatisticsF
{
   public class UDFHelper
    {
        [ExcelFunction(Name = "DOWNSIDEDEVIATION", Category = "Financial Statistics", Description = "This measure is similar to the loss standard deviation except the downside deviation considers only returns that fall below a defined minimum acceptable return (MAR)", HelpTopic = "This measure is similar to the loss standard deviation except the downside deviation considers only returns that fall below a defined minimum acceptable return (MAR) rather than the arithmetic mean. For example, if the MAR is 7%, the downside deviation would measure the variation of each period that falls below 7%. (The loss standard deviation, on the other hand, would take only losing periods, calculate an average return for the losing periods, and then measure the variation between each losing return and the losing return average).")]
        public static double DownsideDeviation(double[] range, double minimalAcceptableReturn)
        {
          return  MathNet.Numerics.Financial.AbsoluteRiskMeasures.DownsideDeviation(range, minimalAcceptableReturn);
        }
        [ExcelFunction(Name = "GAINLOSSRATIO", Category = "Financial Statistics", Description = "Measures a fund’s average gain in a gain period divided by the fund’s average loss in a losing period.Periods can be monthly or quarterly depending on the data frequency.", HelpTopic = "Measures a fund’s average gain in a gain period divided by the fund’s average loss in a losing period. Periods can be monthly or quarterly depending on the data frequency.")]
        public static double GainLossRatio(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteRiskMeasures.GainLossRatio(range);
        }
        [ExcelFunction(Name = "GAINSTANDARDDEVIATION", Category = "Financial Statistics", Description = "Calculation is similar to Standard Deviation , except it calculates an average (mean) return only for periods with a gain and measures the variation of only the gain periods around the gain mean. Measures the volatility of upside performance. © Copyright 1996, 1999 Gary L.Gastineau. First Edition. © 1992 Swiss Bank Corporation.", HelpTopic = "Calculation is similar to Standard Deviation , except it calculates an average (mean) return only for periods with a gain and measures the variation of only the gain periods around the gain mean. Measures the volatility of upside performance. © Copyright 1996, 1999 Gary L.Gastineau. First Edition. © 1992 Swiss Bank Corporation.")]
        public static double GainStandardDeviation(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteRiskMeasures.GainStandardDeviation(range);
        }
        [ExcelFunction(Name = "LOSSSTANDARDDEVIATION", Category = "Financial Statistics", Description = "Calculation is similar to Standard Deviation , except it calculates an average (mean) return only for periods with a gain and measures the variation of only the gain periods around the gain mean. Measures the volatility of upside performance. © Copyright 1996, 1999 Gary L.Gastineau. First Edition. © 1992 Swiss Bank Corporation.", HelpTopic = "Similar to standard deviation, except this statistic calculates an average(mean) return for only the periods with a loss and then measures the variation of only the losing periods around this loss mean. This statistic measures the volatility of downside performance.")]

        public static double LossStandardDeviation(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteRiskMeasures.LossStandardDeviation(range);
        }

       
        [ExcelFunction(Name = "SEMIDEVIATION", Category = "Financial Statistics", Description = "A measure of volatility in returns below the mean. It's similar to standard deviation, but it only looks at periods where the investment return was less than average return.", HelpTopic = "A measure of volatility in returns below the mean. It's similar to standard deviation, but it only looks at periods where the investment return was less than average return.")]
        public static double SemiDeviation(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteRiskMeasures.SemiDeviation(range);
        }
        
        [ExcelFunction(Name = "COMPOUNDRETURN", Category = "Financial Statistics", Description = "Compound Monthly Return or Geometric Return or Annualized Return",  HelpTopic = "Compound Monthly Return or Geometric Return or Annualized Return")]
        public static double CompoundReturn(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteReturnMeasures.CompoundReturn(range);
        }
        
        [ExcelFunction(Name = "GAINMEAN", Category = "Financial Statistics", Description = "Average Gain or Gain Mean This is a simple average (arithmetic mean) of the periods with a gain. It is calculated by summing the returns for gain periods (return 0) and then dividing the total by the number of gain periods.", HelpTopic = "Average Gain or Gain Mean This is a simple average (arithmetic mean) of the periods with a gain. It is calculated by summing the returns for gain periods (return 0) and then dividing the total by the number of gain periods.")]
        public static double GainMean(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteReturnMeasures.GainMean(range);
        }
        
        [ExcelFunction(Name = "LOSSMEAN", Category = "Financial Statistics", Description = "Average Loss or LossMean This is a simple average (arithmetic mean) of the periods with a loss. It is calculated by summing the returns for loss periods (return &lt; 0) and then dividing the total by the number of loss periods.", HelpTopic = "Average Loss or LossMean This is a simple average (arithmetic mean) of the periods with a loss. It is calculated by summing the returns for loss periods (return &lt; 0) and then dividing the total by the number of loss periods.")]
        public static double LossMean(double[] range)
        {
            return MathNet.Numerics.Financial.AbsoluteReturnMeasures.LossMean(range);
        }

        [ExcelFunction(Name = "COVARIANCE", Category = "Statistics",
        Description = "Estimates the unbiased population covariance from the provided two sample arrays. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.",
            HelpTopic = "Estimates the unbiased population covariance from the provided two sample arrays. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.")]
        public static double Covariance(double[] samples1, double[] samples2)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.Covariance(samples1, samples2);
        }
        [ExcelFunction(Name = "FIVENUMBERSUMMARYINPLACE", Category = "Statistics",
       Description = "Estimates {min, lower-quantile, median, upper-quantile, max} from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
           HelpTopic = "Estimates {min, lower-quantile, median, upper-quantile, max} from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double[] FiveNumberSummaryInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.FiveNumberSummaryInplace(range);
        }
        [ExcelFunction(Name = "GEOMETRICMEAN", Category = "Statistics",
       Description = "Evaluates the geometric mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Evaluates the geometric mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double GeometricMean(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.GeometricMean(range);
        }
        [ExcelFunction(Name = "HARMONICMEAN", Category = "Statistics",
       Description = "Evaluates the harmonic mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Evaluates the harmonic mean of the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double HarmonicMean(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.HarmonicMean(range);
        }

        [ExcelFunction(Name = "INTERQUARTILERANGEINPLACE", Category = "Statistics",
       Description = "Estimates the inter-quartile range from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
           HelpTopic = "Estimates the inter-quartile range from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double InterquartileRangeInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.InterquartileRangeInplace(range);
        }

        [ExcelFunction(Name = "LOWERQUARTILEINPLACE", Category = "Financial Statistics",
       Description = "Estimates the first quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
           HelpTopic = "Estimates the first quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double LowerQuartileInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.LowerQuartileInplace(range);
        }
    [ExcelFunction(Name = "MAXIMUM", Category = "Statistics",
       Description = "Returns the largest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Returns the largest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double Maximum(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.Maximum(range);
        }

        [ExcelFunction(Name = "MAXIMUMABSOLUTE", Category = "Financial Statistics",
       Description = "Returns the largest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Returns the largest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double MaximumAbsolute(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.MaximumAbsolute(range);
        }

        [ExcelFunction(Name = "MEAN", Category = "Statistics",
       Description = "Estimates the arithmetic sample mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Estimates the arithmetic sample mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double Mean(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.Mean(range);
        }
    [ExcelFunction(Name = "MEANSTANDARDDEVIATION", Category = "Statistics",
       Description = "Estimates the arithmetic sample mean and the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for standard deviation if data has less than two entries or if any entry is NaN.",
           HelpTopic = "Estimates the arithmetic sample mean and the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for standard deviation if data has less than two entries or if any entry is NaN.")]
        public static double[] MeanStandardDeviation(double[] range)
        {
            Tuple<double, double> ret = MathNet.Numerics.Statistics.ArrayStatistics.MeanStandardDeviation(range);
            return new double[] { ret.Item1, ret.Item2 };
        }
[ExcelFunction(Name = "MEANVARIANCE", Category = "Statistics",
       Description = "Estimates the arithmetic sample mean and the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for variance if data has less than two entries or if any entry is NaN.",
           HelpTopic = "Estimates the arithmetic sample mean and the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN for mean if data is empty or any entry is NaN and NaN for variance if data has less than two entries or if any entry is NaN.")]
        public static double[,]  MeanVariance(double[] range)
        {
            Tuple<double, double> ret= MathNet.Numerics.Statistics.ArrayStatistics.MeanVariance(range);
            
           return new double[,] { { ret.Item1, ret.Item2 } };
         
        }
       
        [ExcelFunction(Name = "MEDIANINPLACE", Category = "Statistics",
       Description = "Estimates the median value from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.",
           HelpTopic = "Estimates the median value from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double MedianInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.MedianInplace(range);
        }
        [ExcelFunction(Name = "MINIMUM", Category = "Statistics",
       Description = "Returns the smallest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Returns the smallest value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double Minimum(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.Minimum(range);
        }
        [ExcelFunction(Name = "MINIMUMABSOLUTE", Category = "Statistics",
       Description = "Returns the smallest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
           HelpTopic = "Returns the smallest absolute value from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double MinimumAbsolute(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.MinimumAbsolute(range);
        }
        [ExcelFunction(Name = "ORDERSTATISTICINPLACE", Category = "Statistics",
       Description = "Returns the order statistic (order 1..N) from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.",
           HelpTopic = "Returns the order statistic (order 1..N) from the unsorted data array. WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double OrderStatisticInplace(double[] range, int order)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.OrderStatisticInplace(range,order);
        }

        [ExcelFunction(Name = "PERCENTILEINPLACE", Category = "Statistics",
      Description = "Estimates the p-Percentile value from the unsorted data array. If a non-integer Percentile is needed, use Quantile instead. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
          HelpTopic = "Estimates the p-Percentile value from the unsorted data array. If a non-integer Percentile is needed, use Quantile instead. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double PercentileInplace(double[] range, int p)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.PercentileInplace(range, p);
        }
    [ExcelFunction(Name = "POPULATIONCOVARIANCE", Category = "Statistics",
    Description = "Evaluates the population covariance from the full population provided as two arrays. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.",
        HelpTopic = "Evaluates the population covariance from the full population provided as two arrays. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.")]
        public static double PopulationCovariance(double[] population1, double[] population2)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.PopulationCovariance(population1, population2);
        }
[ExcelFunction(Name = "POPULATIONSTANDARDDEVIATION", Category = "Statistics",
    Description = "Evaluates the population standard deviation from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.",
        HelpTopic = "Evaluates the population standard deviation from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.")]
        public static double PopulationStandardDeviation(double[] population)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.PopulationStandardDeviation(population);
        }

      [ExcelFunction(Name = "POPULATIONVARIANCE", Category = "Statistics",
      Description = "Evaluates the population variance from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.",
          HelpTopic = "Evaluates the population variance from the full population provided as unsorted array. On a dataset of size N will use an N normalizer and would thus be biased if applied to a subset. Returns NaN if data is empty or if any entry is NaN.")]
        public static double PopulationVariance(double[] population)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.PopulationVariance(population);
        }

      [ExcelFunction(Name = "QUANTILECUSTOMINPLACE", Category = "Statistics",
      Description = "Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. The quantile definition can be specified by 4 parameters a, b, c and d, consistent with Mathematica. WARNING: Works inplace and can thus causes the data array to be reordered.",
          HelpTopic = "Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. The quantile definition can be specified by 4 parameters a, b, c and d, consistent with Mathematica. WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double QuantileCustomInplace(double[] range, double tau, double a, double b, double c, double d)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.QuantileCustomInplace(range,tau,a,b,c,d);
        }

      [ExcelFunction(Name = "QUANTILEINPLACE", Category = "Statistics",
      Description = "Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
          HelpTopic = "Estimates the tau-th quantile from the unsorted data array. The tau-th quantile is the data value where the cumulative distribution function crosses tau. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double QuantileInplace(double[] range, double tau)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.QuantileInplace(range, tau);
        }

      [ExcelFunction(Name = "RANKSINPLACE", Category = "Statistics",
      Description = "Evaluates the rank of each entry of the unsorted data array. The rank definition can be specified to be compatible with an existing system. WARNING: Works inplace and can thus causes the data array to be reordered.",
          HelpTopic = "Evaluates the rank of each entry of the unsorted data array. The rank definition can be specified to be compatible with an existing system. WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double[] RanksInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.RanksInplace(range);
        }

      [ExcelFunction(Name = "ROOTMEANSQUARE", Category = "Statistics",
      Description = "Estimates the root mean square (RMS) also known as quadratic mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.",
          HelpTopic = "Estimates the root mean square (RMS) also known as quadratic mean from the unsorted data array. Returns NaN if data is empty or any entry is NaN.")]
        public static double RootMeanSquare(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.RootMeanSquare(range);
        }

      [ExcelFunction(Name = "STANDARDDEVIATION", Category = "Statistics",
      Description = "Estimates the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.",
          HelpTopic = "Estimates the unbiased population standard deviation from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.")]
        public static double StandardDeviation(double[] range)

        {
            return MathNet.Numerics.Statistics.ArrayStatistics.StandardDeviation(range);
        }

      [ExcelFunction(Name = "UPPERQUARTILEINPLACE", Category = "Statistics",
      Description = "Estimates the third quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.",
          HelpTopic = "Estimates the third quartile value from the unsorted data array. Approximately median-unbiased regardless of the sample distribution (R8). WARNING: Works inplace and can thus causes the data array to be reordered.")]
        public static double UpperQuartileInplace(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.UpperQuartileInplace(range);
        }
        [ExcelFunction(Name = "VARIANCE", Category = "Statistics",
       Description = "Estimates the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.",
           HelpTopic = "Estimates the unbiased population variance from the provided samples as unsorted array. On a dataset of size N will use an N-1 normalizer (Bessel's correction). Returns NaN if data has less than two entries or if any entry is NaN.")]
        public static double Variance(double[] range)
        {
            return MathNet.Numerics.Statistics.ArrayStatistics.Variance(range);
        }


        /* [ExcelFunction(Name = "LOSSSTANDARDDEVIATION", Category = "Statistics",
       Description = "",
           HelpTopic = "")] */
    }
}