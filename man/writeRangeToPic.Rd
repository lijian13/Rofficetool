\name{writeRangeToPic}
\alias{writeRangeToPic}
\title{Write a data.frame to a excel range.}
\usage{
writeRangeToPic(excelfile, shtindex, rangestring, picfile, chartsize = c(800,
  600))
}
\arguments{
  \item{df}{the data.frame}

  \item{excelfile}{path of a excel file}

  \item{shtindex}{sheet index}

  \item{rangelefttop}{the range string, such as "A1"}
}
\value{
TRUE or FALSE
}
\description{
Write a data.frame to a excel range.
}
\author{
Jian Li <\email{rweibo@sina.com}>
}

