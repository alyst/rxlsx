# Automate the making of the package
#
#


##################################################################
#
.update.DESCRIPTION <- function(packagedir, version)
{
  file <- paste(packagedir, "DESCRIPTION", sep="") 
  DD  <- readLines(file)
  ind  <- grep("Version: ", DD)
  aux <- strsplit(DD[ind], " ")[[1]]
  
  if (is.null(version)){   # increase by one 
    vSplit    <- strsplit(aux[2], "\\.")[[1]]
    vSplit[3] <- as.character(as.numeric(vSplit[3])+1) 
    version <- paste(vSplit, sep="", collapse=".")
  }   
  DD[ind] <- paste(aux[1], version)

  ind <- grep("Date: ", DD)
  aux <- strsplit(DD[ind], " ")[[1]]
  DD[ind] <- paste(aux[1], Sys.Date())
  
  writeLines(DD, con=file)
  return(version)
}

##################################################################
#
.build.java <- function()
{
  # build maven project
  # and put the jars (package and its dependencies) to the inst/java/ directory
  system(paste("mvn -f ", file.path(pkgdir,'other', 'pom.xml'), ' package' ) )

  ## # move the source files to have for reference ... 
  ## file.copy("src/dev/RInterface.java", paste(pkgdir, 
  ##   "other/RInterface.java", sep=""), overwrite=TRUE)
  ## file.copy("src/tests/TestRInterface.java", paste(pkgdir, 
  ##   "other/TestRInterface.java", sep=""), overwrite=TRUE)
  invisible()
}

##################################################################
#
.setEnv <- function(computer=c("HOME", "LAPTOP", "WORK"))
{
  if (computer=="WORK") {
    pkgdir  <<- "C:/google/rexcel/trunk/"
    outdir  <<- "H:/"
    Rcmd    <<- "S:/All/Risk/Software/R/R-2.12.1/bin/i386/Rcmd"
    javadir <<- "C:/Documents and Settings/e47187/workspace/xlsx/"
  } else if (computer == "LAPTOP") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp/"
    Rcmd      <<- "R CMD"
    javadir   <<- "/home/adrian/workspace/xlsx/"
  } else if (computer == "HOME") {
    pkgdir    <<- "/home/adrian/Documents/rexcel/trunk/"
    outdir    <<- "/tmp"
    Rcmd      <<- "R CMD"
    javadir   <<- "/home/adrian/workspace/xlsx/"
  } else if (computer == "WORK2") {  
    pkgdir  <<- "C:/google/rexcel/trunk/"
    outdir  <<- "H:/"
    Rcmd    <<- '"C:/Program Files/R/R-2.14.1/bin/i386/Rcmd"'
    javadir <<- "C:/Documents and Settings/e47187/workspace/xlsx/"
  } else {
  }

  invisible()
}

##################################################################
##################################################################

#version <- NULL        # keep increasing the minor
version <- "0.5.0"      # if you want to set it by hand

.setEnv("WORK")   # "HOME" "WORK2" "LAPTOP"

.build.java()  # move java classes

# change the version
version <- .update.DESCRIPTION(pkgdir, version)

# make the package
setwd(outdir)
cmd <- paste(Rcmd, "build --force", pkgdir)
print(cmd)

system(cmd)

package.gz <- paste("xlsx_",version, ".tar.gz", sep="")
install.packages(package.gz, repos=NULL, type="source")


# do you pass all my tests?! Open another R session ... 
cat(paste("require(xlsx); source('", pkgdir,
          "other/runUnitTests.R')", sep=""), "\n\n")


# make the package for CRAN
cmd <- paste(Rcmd, "build", pkgdir)
print(cmd); system(cmd)


# check source with --as-cran on the tarball before submitting it
cmd <- paste(Rcmd, "check --as-cran", package.gz)
print(cmd); system(cmd)




## .deepCopy("C:/Users/adrian/R/findataweb/temp/xlsx/trunk/",
##    "C:/Users/adrian/R/findataweb/temp/xlsx/tags/0.1.3/")


## ##################################################################
## # Copy one folder to another without the .svn dirs
## #  fromDir <- "C:/Users/adrian/R/findataweb/temp/xlsx/trunk"
## #  toDir <- "C:/Temporary/Downloads/xlsx"
## #  .deepCopy(fromDir, toDir)
## #
## .deepCopy <- function(fromDir, toDir)
## {
##   if (file.info(fromDir)$isdir){
##     fromFiles <- list.files(fromDir, full.names=TRUE)
    
##     for (f in fromFiles){
##       if (file.info(f)$isdir){
##         toDir2 <- paste(toDir, basename(f), sep="/")
##         dir.create(toDir2)
##         .deepCopy(f, toDir2)      
##       } else {
##         file.copy(f, toDir)
##       }
##     }
##   } else {
##     file.copy(fromDir, toDir)
##   }
## }  
