
##' Write a text to a ppt object.
##' 
##' @title Write a text to a ppt object.
##' @param Text The text to be written to ppt.
##' @param pptobject A ppt object created by \code{\link{readPPT}}.
##' @param ipage The page number.
##' @param ishape The shape number.
##' @return invisible TRUE or FALSE.
##' @author Jian Li <\email{rweibo@@sina.com}>

writeTextToPPT <- function(Text, pptobject, ipage = 1, ishape = 1) {

	pptslide <- pptobject[["Com"]][["Slides"]]$Item(ipage)
	pptshape <- pptslide[["Shapes"]]$Item(ishape)
	
	tmp.text <- pptshape[["TextFrame"]][["TextRange"]]
	tmp.text[["Text"]] <- Text
	
	invisible(TRUE)
}

