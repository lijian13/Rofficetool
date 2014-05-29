
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

writeDfToPPT <- function(dF, pptfile, ipage = 1, ishape = 1) {

	pptfile <- normalizePath(pptfile, winslash = "/", mustWork = TRUE)
	pptapp <- COMCreate("PowerPoint.Application")
	pptpre <- pptapp[["Presentations"]]$Open(pptfile)
	pptslide <- pptpre[["Slides"]]$Item(ipage)
	pptshape <- pptslide[["Shapes"]]$Item(ishape)
	if (pptshape$type() != 19) stop("This shape is not a table!")
	ntblrow <- pptshape[["Table"]][["Rows"]]$Count()
	ntblcol <- pptshape[["Table"]][["Columns"]]$Count()
	
	for (i in 1:ntblrow) {
		for (j in 1:ntblcol) {
			if (!is.null(dF[i,j]) && !is.na(dF[i,j])) {
				tmp.table <- pptshape[["Table"]]$Cell(i, j)[["Shape"]][["TextFrame"]][["TextRange"]]
				tmp.table[["Text"]] <- dF[i, j]
			}
		}
	}
	
	pptpre$Save()
	pptpre$Close()
}

