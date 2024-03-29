\name{readColumns}
\alias{readColumns}
\title{ Read a contiguous set of columns from sheet into an R data.frame}
\description{
  Read a contiguous set of columns from sheet into an R data.frame.
  Uses the \code{RInterface} for speed.   
}
\usage{
readColumns(sheet, startColumn, endColumn, startRow,
  endRow=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA,
  ...)
}
\arguments{
  \item{sheet}{a \code{\link{Worksheet}} object.}

  \item{startColumn}{a numeric value for the starting column.}

  \item{endColumn}{a numeric value for the ending column.}

  \item{startRow}{a numeric value for the starting row.}
  
  \item{endRow}{a numeric value for the ending row.  If \code{NULL} it
    reads all the rows in the sheet.  If you request more than the
    existing rows in the sheet, the result will be truncated by the
    actual row number.}
  
  \item{as.data.frame}{a logical value indicating if the result should
    be coerced into a \code{data.frame}.  If \code{FALSE}, the result is
    a list with one element for each column.}
  
  \item{header}{a logical value indicating whether the first row
    corresponding to the first element of the \code{rowIndex} vector
    contains the names of the variables.}
  
  \item{colClasses}{a character vector that represent the class of each
    column.  Recycled as necessary, or if \code{NA} an attempt is made
    to guess the type of each column by reading the first row of data.}
    
  \item{\ldots}{other arguments to \code{data.frame}, for example
    \code{stringsAsFactors}} 
}

\details{
  
Use the \code{readColumns} function when you want to read a rectangular
block of data from an Excel worksheet.  If you request columns which
are blank, these will be read in as empty character "" columns.
Internally, the loop over columns is done in R, the loop over rows is
done in Java, so this function achieves good performance when number of
rows >> number of columns. 

Excel internally stores dates and datetimes as numeric values, and does
not keep track of time zones and DST.  When a numeric column is
formatted as a datetime, it will be converted into \code{POSIXct} class
with a \emph{GMT} timezone.  If you need a \code{Date} column, you need
to specify explicitly using \code{colClasses} argument.

For a numeric column Excels's errors and blank cells will be returned as
NaN values.  Excel's \code{#N/A} will be returned as NA.  Formulas will be
evaluated. For a chracter column, blank cells will be returned as "". 

}
\value{
  A data.frame or a list, depending on the \code{as.data.frame}
  argument. 
}
\author{ Adrian Dragulescu }
\seealso{\code{\link{read.xlsx2}} for reading entire sheets.
  See also \code{\link{addDataFrame}} for writing a \code{data.frame} to a
  sheet.}
\examples{
\dontrun{

  file <- system.file("tests", "test_import.xlsx", package = "xlsx")

  wb     <- loadWorkbook(file)
  sheets <- getSheets(wb)

  sheet <- sheets[["all"]]
  res <- readColumns(sheet, startColumn=3, endColumn=10, startRow=3,
    endRow=7)

  sheet <- sheets[["NAs"]]
  res <- readColumns(sheet, 1, 6, 1,  colClasses=c("Date", "character",
    "integer", rep("numeric", 2),  "POSIXct"))
 

}
}

