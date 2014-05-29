
##' Write a data.frame to a excel range.
##' 
##' @title Write a data.frame to a excel range.
##' @param df the data.frame
##' @param excelfile path of a excel file
##' @param shtindex sheet index
##' @param rangelefttop the range string, such as "A1"
##' @return TRUE or FALSE
##' @author Jian Li <\email{rweibo@@sina.com}>


#dF <- tdf2
#excelfile <- "E:/Mango/Training/Youku/20140515/examples/report/table_1.xlsx"
#shtindex <- 1
#rangelefttop <- c(5, 2)
# writeDfToRange( tdf2, "E:/Mango/Training/Youku/20140515/examples/report/table_1.xlsx", 1, c(5,2))

writePicToPPT <- function(pptfile, picfile, ipage = 1, 
	Position = c(0, 100), Size = c(300, 200)) {
    pptfile <- normalizePath(pptfile, winslash = "/", mustWork = TRUE)
	picfile <- normalizePath(picfile, winslash = "\\", mustWork = FALSE)

	pptapp <- COMCreate("PowerPoint.Application")
	pptpre <- pptapp[["Presentations"]]$Open(pptfile)
	pptslide <- pptpre[["Slides"]]$Item(ipage)
	
	pptshapes <- pptslide[["Shapes"]]
	pptshapes$AddPicture(picfile,0,-1,Position[1],Position[2],Size[1],Size[2])
	pptpre$Save()
	pptpre$Close()
}

