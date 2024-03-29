\name{addDataFrame}
\alias{addDataFrame}
\title{Add a \code{data.frame} to a sheet.}
\description{
  Add a \code{data.frame} to a sheet, allowing for different column styles.
  Useful when constructing the spreadsheet from scratch. 
}
\usage{

addDataFrame(x, sheet, col.names=TRUE, row.names=TRUE,
  startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
  rownamesStyle=NULL, showNA=FALSE, characterNA="")

}
\arguments{
  \item{x}{a \code{data.frame}.}
  
  \item{sheet}{a \code{\link{Sheet}} object.}
  
  \item{col.names}{a logical value indicating if the column names of
    \code{x} are to be written along with \code{x} to the file.}
  
  \item{row.names}{a logical value indicating whether the row names of 
    \code{x} are to be written along with \code{x} to the file.}
  
  \item{startRow}{a numeric value for the starting row.}
  
  \item{startColumn}{a numeric value for the starting column.}
  
  \item{colStyle}{a list of \code{\link{CellStyle}}.  If the name of the
     list element is the column number, it will be used to set the style
     of the column.  Columns of type \code{Date} and \code{POSIXct} are
     styled automatically even if \code{colSyle=NULL}.}
   
  \item{colnamesStyle}{a \code{\link{CellStyle}} object to customize the
    table header.}
  
  \item{rownamesStyle}{a \code{\link{CellStyle}} object to customize the
    row names (if \code{row.names=TRUE}).}
  
  \item{showNA}{a boolean value to control how NA's are displayed on the
    sheet.  If \code{FALSE}, NA values will be represented as blank
    cells.}
  
  \item{characterNA}{a string value to control how character NA will be
    shown in the spreadsheet.}
}

\details{
  Internally this function uses Java internally for looping and the
  result is a significant improvement in performance compared with a
  cell by cell application of \code{\link{setCellValue}}. 
  
  It is difficult to treat \code{NA}'s consistently between R
  and Excel via Java.  Most likely, users of Excel will want to see
  \code{NA}'s  as blank cells.  In R character \code{NA}'s are simply
  characters, which for Excel means "NA".
}

\value{
  None.  The modification to the workbook is done in place.
}
\author{Adrian Dragulescu}
\examples{

  wb <- createWorkbook()
  sheet  <- createSheet(wb, sheetName="addDataFrame1")
  data <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("1999-01-01"), by="1 year", length.out=10),
    bool=c(TRUE, FALSE), log=log(1:10),
    rnorm=10000*rnorm(10),
    datetime=seq(as.POSIXct("2011-11-06 00:00:00", tz="GMT"), by="1 hour",
      length.out=10))
  cs1 <- CellStyle(wb) + Font(wb, isItalic=TRUE)           # rowcolumns
  cs2 <- CellStyle(wb) + Font(wb, color="blue")
  cs3 <- CellStyle(wb) + Font(wb, isBold=TRUE) + Border()  # header
  addDataFrame(data, sheet, startRow=3, startColumn=2, colnamesStyle=cs3,
    rownamesStyle=cs1, colStyle=list(`2`=cs2, `3`=cs2))


  # Don't forget to save the workbook ...  
  # saveWorkbook(wb, file) 
}

