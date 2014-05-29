\name{writePicToPPT}
\alias{writePicToPPT}
\title{Write a data.frame to a excel range.}
\usage{
writePicToPPT(pptfile, picfile, ipage = 1, Position = c(0, 100),
  Size = c(300, 200))
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

