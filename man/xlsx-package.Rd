\name{xlsx-package}
\alias{xlsx-package}
\alias{xlsx}
\docType{package}
\title{
  Read, write, format Excel 2007 and Excel 97/2000/XP/2003 files
}
\description{
  
  The \code{xlsx} package gives programatic control of Excel files using
  R.  A high level API allows the user to read a sheet of an xlsx
  document into a \code{data.frame} and write a \code{data.frame} to a
  file.  Lower level functionality permits the direct manipulation of
  sheets, rows and cells.  For example, the user has control to set
  colors, fonts, data formats, add borders, hide/unhide sheets,
  add/remove rows, add/remove sheets, etc.
  
  Behind the scenes, the \code{xlsx} package uses a java library from
  the Apache project, \url{http://poi.apache.org/index.html}.  This
  Apache project provides a Java API to Microsoft Documents (Excel,
  Word, PowerPoint, Outlook, Visio, etc.)  By using the \code{rJava}
  package that links \R and Java, we can piggyback on the excellent work
  already done by the folks at the Apache project and provide this
  functionality in R.  The \code{xlsx} package uses only a subset of the
  Apache POI project, namely the one dealing with Excel files.  All the
  necessary jar files are kept in package \code{xlsxjars} that is
  imported by package \code{xlsx}.
  
  A collection of tests that can be used as examples are located 
  in folder \code{/tests/}.  They are a good source of examples of how
  to use the package. 

  Please see \url{http://code.google.com/p/rexcel/} for Wiki and the
  development version.  Issues and questions should be reported using
  the dedicated mailing list 
  \url{http://groups.google.com/group/R-package-xlsx}.
  
}
\details{
\tabular{ll}{
Package: \tab xlsx\cr
Type: \tab Package\cr
Version: \tab 0.4.3\cr
Date: \tab 2012-xx-xx\cr
License: \tab GPL-3\cr
}

}
\author{
Adrian A. Dragulescu

Maintainer: Adrian A. Dragulescu <adrian.dragulescu@gmail.com>

}
\references{
  Apache POI project for Microsoft Excel format:
  \url{http://poi.apache.org/spreadsheet/index.html}.
  
  The Java Doc detailing the classes:
  \url{http://poi.apache.org/apidocs/index.html}.  This can be useful
  if you are looking for something that is not exposed in R as it may be
  available on the Java side.  Inspecting the source code for some the
  the \R functions in this package can show you how to do it (even if
  you are Java shy). 
}
\keyword{ package }
\seealso{\code{\link{Workbook}} for ways to work with \code{Workbook}
  objects.}
\examples{
\dontrun{

require(xlsx)

# example of reading xlsx sheets
file <- system.file("tests", "test_import.xlsx", package = "xlsx")
res <- read.xlsx(file, 2)  # read the second sheet

# example of writing xlsx sheets
file <- tempfile()
write.xlsx(USArrests, file=file)

}
}
