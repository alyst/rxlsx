# Write a data.frame to a new xlsx file. 
# with a java back-end
#

write.xlsx2 <- function(x, file, sheetName="Sheet1",
  col.names=TRUE, row.names=TRUE, append=FALSE, ...)
{
  if (append){
    wb <- loadWorkbook(file)
  } else {
    wb <- createWorkbook()
  }  
  sheet <- createSheet(wb, sheetName)

  addDataFrame(x, sheet, col.names=col.names, row.names=row.names,
    startRow=1, startColumn=1, colStyle=NULL, colnamesStyle=NULL,
    rownamesStyle=NULL)

  saveWorkbook(wb, file)  
  
  invisible()
}


