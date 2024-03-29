# Add a data.frame to a sheet
#
# colStyle can be a structure of CellStyle with names representing column
#  index.  
#
# I shouldn't offer to change the data for the user, with characterNA="",
# they should do it themselves.  It's not that hard!
#
addDataFrame <- function(x, sheet, col.names=TRUE, row.names=TRUE,
  startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
  rownamesStyle=NULL, showNA=FALSE, characterNA="", transpose = FALSE)
{
  if (!is.data.frame(x))
    x <- data.frame(x)    # just because the error message is too ugly

  if (row.names) {        # add rownames to data x                   
    x <- cbind(rownames=rownames(x), x)
    if (!is.null(colStyle))
      names(colStyle) <- as.numeric(names(colStyle)) + 1
  }
  
  wb <- sheet$getWorkbook()
  classes <- unlist(sapply(x, class))
  if ("Date" %in% classes) 
    csDate <- CellStyle(wb) + DataFormat("m/d/yyyy")
  if ("POSIXct" %in% classes) 
    csDateTime <- CellStyle(wb) + DataFormat("m/d/yyyy h:mm:ss;@")

  iOffset <- if (col.names) 1L else 0L
  jOffset <- if (row.names) 1L else 0L
  indX1   <- as.integer(startRow-1L)        # index of top row
  indY1   <- as.integer(startColumn-1L)     # index of top column
  
  # create a new interface object 
  cellBlock <- .jnew("org/cran/rxlsx/RCellBlock", sheet, indX1, indY1,
                 nrow(x) + iOffset, ncol(x),
                 transpose, TRUE)
  
  if (col.names) {                   # insert colnames
    .jcall( cellBlock, "V", "setRowData",
            0L, jOffset, .jarray(if (row.names) names(x)[-1] else names(x)),
            showNA, if ( !is.null(colnamesStyle) ) colnamesStyle$ref else .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }
  # insert one column at a time, and style it if it has style
  # Dates and POSIXct columns get styled if not overridden. 
  for (j in 1:ncol(x)) {
    colStyle <-
      if ((j==1) && (row.names) && (!is.null(rownamesStyle))) {
        rownamesStyle
      } else if (as.character(j) %in% names(colStyle)) {
        colStyle[[as.character(j)]]
      } else if ("Date" %in% class(x[,j])) {
        csDate
      } else if ("POSIXt" %in% class(x[,j])) {
        csDateTime
      } else {
        NULL
      }
#browser()
    xj <- x[,j]
    if ("integer" %in% class(xj)) {
      aux <- xj
    } else if (any(c("numeric", "Date", "POSIXt") %in% class(xj))) {
      aux <- if ("Date" %in% class(xj)) {
          as.numeric(xj)+25569
        } else if ("POSIXt" %in% class(x[,j])) {
          as.numeric(xj)/86400 + 25569
        } else {
          xj
        }
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- NaN          # encode the numeric NAs as NaN for java
    } else {
      aux <- as.character(x[,j])
      haveNA <- is.na(aux)
      if (any(haveNA))
        aux[haveNA] <- characterNA
    }
    .jcall( cellBlock, "V", "setColData", as.integer(j+jOffset-1L), iOffset, .jarray(aux),
            showNA, if ( !is.null(colStyle) ) colStyle$ref else .jnull('org/apache/poi/ss/usermodel/CellStyle') )
  }
  
  return ( cellBlock )
}

setCellStyle <- function( cellBlock, style )
{
    cellBlock$setCellStyle( style )
    invisible()
}

setCellStyle <- function( cellBlock, style, rowIndex = NULL, colIndex = NULL )
{
    cellBlock$setCellStyle( style, .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

setFill <- function( cellBlock, fill, rowIndex = NULL, colIndex = NULL )
{
    .jcall( cellBlock, 'V', 'setFill',
            .xssfcolor( fill$foregroundColor ), .xssfcolor( fill$backgroundColor ),
            .jshort(FILL_STYLES_[[fill$pattern]]),
            .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

setFont <- function( cellBlock, font, rowIndex = NULL, colIndex = NULL )
{
    cellBlock$setFont( font$ref, .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

putBorder <- function( cellBlock, border, rowIndex = NULL, colIndex = NULL )
{
    border_none <- BORDER_STYLES_[['BORDER_NONE']]
    borders <- c( TOP = border_none, BOTTOM = border_none,
                  LEFT = border_none, RIGHT = border_none )
    null_color <- .jnull('org/apache/poi/xssf/usermodel/XSSFColor')
    border_colors <- c( TOP = null_color, BOTTOM = null_color,
                        LEFT = null_color, RIGHT = null_color )
    borders[ border$position ] <- sapply( border$pen, function( pen ) BORDER_STYLES_[pen] )
    border_colors[ border$position ] <- sapply( border$color, .xssfcolor )
    
    .jcall( cellBlock, "V", "putBorder",
            .jshort(borders[['TOP']]), border_colors[['TOP']],
            .jshort(borders[['BOTTOM']]), border_colors[['BOTTOM']],
            .jshort(borders[['LEFT']]), border_colors[['LEFT']],
            .jshort(borders[['RIGHT']]), border_colors[['RIGHT']],
            .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

getCell <- function( cellBlock, rowIndex, colIndex )
{
    return ( .jcall( cellBlock, 'Lorg/apache/poi/ss/usermodel/Cell;', 'getCell',
              rowIndex - 1L, colIndex - 1L ) )
}
