
##' Write a data.frame to a ppt object.
##' 
##' @title Write a data.frame to a ppt object.
##' @param Df The data.frame to be written to ppt.
##' @param pptobject A ppt object created by \code{\link{readPPT}}.
##' @param ipage The page number.
##' @param ishape The shape number.
##' @param startrow The start row in the table of ppt.
##' @param startcol The start column in the table of ppt.
##' @return invisible TRUE or FALSE.
##' @author Jian Li <\email{rweibo@@sina.com}>

writeDfToPPT <- function(Df, pptobject, ipage = 1, ishape = 1, startrow = 1, startcol = 1) {

	pptslide <- pptobject[["Com"]][["Slides"]]$Item(ipage)
	pptshape <- pptslide[["Shapes"]]$Item(ishape)
	if (pptshape$type() != 19) stop("This shape is not a table!")
	ntblrow <- pptshape[["Table"]][["Rows"]]$Count()
	ntblcol <- pptshape[["Table"]][["Columns"]]$Count()
	ndfrow <- nrow(Df)
	ndfcol <- ncol(Df)
	if (ntblrow - ndfrow != startrow - 1) stop("Wrong row number!")
	if (ntblcol - ndfcol != startcol - 1) stop("Wrong column number!")
	
	for (i in 1:ndfrow) {
		for (j in 1:ndfcol) {
			if (!is.null(Df[i, j]) && !is.na(Df[i,j])) {
				tmp.table <- pptshape[["Table"]]$Cell(i + startrow - 1, j + startcol - 1)[["Shape"]][["TextFrame"]][["TextRange"]]
				tmp.table[["Text"]] <- Df[i, j]
			}
		}
	}
	
	invisible(TRUE)
}

