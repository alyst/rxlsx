\name{Cell}
\alias{Cell}
\alias{createCell}
\alias{getCells}
\alias{setCellValue}
\alias{getCellValue}
\alias{getMatrixValues}
\title{Functions to manipulate cells.}
\description{
Functions to manipulate cells. 
}
\usage{

createCell(row, colIndex=1:5)

getCells(row, colIndex=NULL, simplify=TRUE)

setCellValue(cell, value, richTextString=FALSE)

getCellValue(cell, keepFormulas=FALSE, encoding="unknown")

getMatrixValues(sheet, rowIndex, colIndex, ...)

}
\arguments{
  \item{row}{a list of row objects. See \code{Row}.}
  \item{rowIndex}{a numeric vector specifying the index of rows.}
  \item{colIndex}{a numeric vector specifying the index of columns.}
  \item{simplify}{a logical value.  If \code{TRUE}, the result will be
    unlisted.}
  \item{value}{an R variable of length one.}
  \item{richTextString}{a logical value indicating if the value should
    be inserted into the Excel cell as rich text.}
  \item{keepFormulas}{a logical value.  If \code{TRUE} the formulas will 
    be returned as characters instead of being explicitly evaluated.}
  \item{encoding}{A character value to set the encoding, for example "UTF-8".}
  \item{cell}{a \code{Cell} object.}
  \item{sheet}{a \code{Sheet} object.}
  \item{\ldots}{other arguments to be passed to \code{getCellValue}.}
}

\value{
  
  \code{createCell} creates a matrix of lists, each element of the list
  being a java object reference to an object of type Cell representing
  an empty cell.  The dimnames of this matrix are taken from the names
  of the rows and the \code{colIndex} variable.
    
  \code{getCells} returns a list of java object references for all the
  cells in the row if \code{colIndex} is \code{NULL}.  If you want to
  extract only a specific columns, set \code{colIndex} to the column 
  indices you are interested.
  
  \code{getCellValue} returns the value in the cell as an R object.
  Type conversions are done behind the scene.  This function is not
  vectorized.
  
  \code{getMatrixValues} is useful to extract blocks of similar cells
  from a spreadsheet.  The result is a matrix.  This is a higher
  function that is often needed.  If the \code{rowIndex} or
  \code{colIndex} exceed the data of the sheet, \code{NA}'s will be
  returned. 

  
}
\details{
  \code{setCellValue} writes the content of an R variable into the cell.
  \code{Date} and \code{POSIXct} objects are passed in as numerical
  values.  To format them as dates in Excel see \code{\link{CellStyle}}.
  This function is not vectorized. 
}

\author{Adrian Dragulescu}
\seealso{To format cells, see \code{\link{CellStyle}}.  For rows see
  \code{\link{Row}}, for sheets see \code{\link{Sheet}}.}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)  
sheets <- getSheets(wb)

sheet <- sheets[['mixedTypes']]      # get second sheet
rows  <- getRows(sheet)   # get all the rows

cells <- getCells(rows)   # returns all non empty cells

values <- lapply(cells, getCellValue) # extract the values

# get a matrix of elements (converted to character)
vv <- getMatrixValues(sheet, 2:3, 1:3)

# write the months of the year in the first column of the spreadsheet
ind <- paste(2:13, ".2", sep="")
mapply(setCellValue, cells[ind], month.name)


}

