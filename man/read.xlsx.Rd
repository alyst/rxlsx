\name{read.xlsx}
\alias{read.xlsx}
\alias{read.xlsx2}
\title{ Read the contents of a worksheet into an R data.frame}
\description{
Read the contents of a worksheet into an R \code{data.frame}.
}
\usage{
read.xlsx(file, sheetIndex, sheetName=NULL, rowIndex=NULL,
  colIndex=NULL, as.data.frame=TRUE, header=TRUE, colClasses=NA,
  keepFormulas=FALSE, encoding="unknown", ...)

read.xlsx2(file, sheetIndex, sheetName=NULL, startRow=1,
  colIndex=NULL, endRow=NULL, as.data.frame=TRUE, header=TRUE,
  colClasses="character", ...)
}
\arguments{ 
  \item{file}{the \emph{absolute} path to the file which the data are to
    be read from.}
  \item{sheetIndex}{a number representing the sheet index in the workbook.}
  \item{sheetName}{a character string with the sheet name.}
  \item{rowIndex}{a numeric vector indicating the rows you want to
    extract.  If \code{NULL}, all rows found will be extracted.}
  \item{colIndex}{a numeric vector indicating the cols you want to
    extract.  If \code{NULL}, all columns found will be extracted.}  
  \item{as.data.frame}{a logical value indicating if the result should
    be coerced into a \code{data.frame}.  If \code{FALSE}, the result is
    a list with one element for each column.}     
  \item{header}{a logical value indicating whether the first row
    corresponding to the first element of the \code{rowIndex} vector
    contains the names of the variables.}
  \item{colClasses}{For \code{read.xlsx} a character vector that
	represent the class of each column.  Recycled as necessary, or if
	the character vector is named, unspecified values are taken to be
	\code{NA}.  For \code{read.xlsx2} A character vector, recycled if
	necessary.  Only \code{numeric} and \code{character} values are
	allowed. }
  \item{keepFormulas}{a logical value indicating if Excel formulas
    should be shown as text in \R and not evaluated before bringing them
    in.}
  \item{encoding}{encoding to be assumed for input strings.  See
	\code{\link[utils]{read.table}}.}
  \item{startRow}{an numeric specifying the index of starting row.}
  \item{endRow}{an numeric specifying the index of the last row to pull.  If
	\code{NULL}, read all the rows in the sheet.}
  \item{\ldots}{other arguments to \code{data.frame}, for example
    \code{stringsAsFactors}} 
}

\details{
  
The \code{read.xlsx} function provides a high level API for reading data
from an Excel worksheet.  It calls several low level functions in
the process.  Its goal is to provide the conveniency of
\code{\link[utils]{read.table}} by borrowing from its signature. 

The function pulls the value of each non empty cell in the worksheet
into a vector of type \code{list} by preserving the data type.  If
\code{as.data.frame=TRUE}, this vector of lists is then formatted into a
rectangular shape.  Special care is needed for worksheets with ragged
data.   

The class type of the variable corresponding to one column in the
worksheet is taken from the class of the first non empty cell in that
column.  If you need to impose a specific class type on a variable, use
the \code{colClasses} argument.  

Excel internally stores dates and datetimes as numeric values, and does
not keep track of time zones and DST.  When a datetime column is 
brought into \R, it is converted to \code{POSIXct} class with a
\emph{GMT} timezone.  Occasional rounding errors may appear and the \R
and Excel string representation my differ by one second.  For
\code{read.xlsx2} bring in a datetime column as a numeric one and then
convert to class \code{POSIXct} or \code{Date}.   

The \code{read.xlsx2} function does more work in Java so it achieves
better performance (an order of magnitude faster on sheets with
100,000 cells or more).  The result of \code{read.xlsx2} will in
general be different from \code{read.xlsx}, because internally
\code{read.xlsx2} uses \code{readColumns} which is tailored for tabular
data.  

}
\value{
  A data.frame or a list, depending on the \code{as.data.frame}
  argument. 
}
\author{ Adrian Dragulescu }
\seealso{\code{\link{write.xlsx}} for writing \code{xlsx} documents.
  See also \code{\link{readColumns}} for reading only a set of columns
  into R.}
\examples{
\dontrun{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")
res <- read.xlsx(file, 1)  # read first sheet
head(res)
#          NA. Population Income Illiteracy Life.Exp Murder HS.Grad Frost   Area
# 1    Alabama       3615   3624        2.1    69.05   15.1    41.3    20  50708
# 2     Alaska        365   6315        1.5    69.31   11.3    66.7   152 566432
# 3    Arizona       2212   4530        1.8    70.55    7.8    58.1    15 113417
# 4   Arkansas       2110   3378        1.9    70.66   10.1    39.9    65  51945
# 5 California      21198   5114        1.1    71.71   10.3    62.6    20 156361
# 6   Colorado       2541   4884        0.7    72.06    6.8    63.9   166 103766
# >


# To convert an Excel datetime colum to POSIXct, do something like:
#   as.POSIXct((x-25569)*86400, tz="GMT", origin="1970-01-01")
# For Dates, use a conversion like:
#   as.Date(x-25569, origin="1970-01-01") 

res2 <- read.xlsx2(file, 1)  

}
}

