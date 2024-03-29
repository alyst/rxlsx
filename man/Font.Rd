\name{Font}
\alias{Font}
\alias{createFont}
\alias{is.Font}
\title{Create a Font object.}
\description{
  Create a Font object.
}
\usage{

Font(wb, color=NULL, heightInPoints=NULL, name=NULL,
  isItalic=FALSE, isStrikeout=FALSE, isBold=FALSE, underline=NULL,
  boldweight=NULL)

createFont(...)

is.Font(x)

}
\arguments{
  \item{wb}{a workbook object as returned by
    \code{\link{createWorkbook}} or \code{\link{loadWorkbook}}.}

  \item{color}{a character specifiying the font color.  Any
    color names as returned by \code{\link[grDevices]{colors}} can be
    used.  For Excel 95 workbooks, only a subset of colors is
    available, see the constant \code{INDEXED_COLORS_}.}
  
  \item{heightInPoints}{a numeric value specifying the font height.
  Usual values are 10, 12, 14, etc.}
  
  \item{name}{a character value for the font to use.  All values that
    you see in Excel should be available, e.g. "Courier New".}

  \item{isItalic}{a logical indicating the font should be italic.}

  \item{isStrikeout}{a logical indicating the font should be stiked out.}

  \item{isBold}{a logical indicating the font should be bold.}

  \item{underline}{a numeric value specifying the thickness of the
  underline.  Allowed values are 0, 1, 2.}

  \item{boldweight}{a numeric value indicating bold weight.  Normal is
    400, regular bold is 700. }

  \item{x}{A Font object, as returned by \code{Font}. }

  \item{\ldots}{arguments get passed to \code{Font}.} 
  
}

\details{
  Default values for \code{NULL} parameters are taken from Excel.  So
  the default font color is black, the default font name is "Calibri", and
  the font height in points is 11.

  \code{createFont} is now deprecated and will be removed in version
  0.5.0 of the package.  
  
  NOTE: You need to have a \code{Workbook} object to attach a
  \code{Font} object to it. 
}
\value{
  \code{Font} returns a list with a java reference to a
  \code{Font} object, and a class attribute "Font".

  \code{is.Font} returns \code{TRUE} if the argument is of class
  "Font" and \code{FALSE} otherwise.

}
\author{Adrian Dragulescu}
\seealso{\code{\link{CellStyle}} for using the a \code{Font} object.}
\examples{
\dontrun{
  font <-  Font(wb, color="blue", isItalic=TRUE)
}
}


