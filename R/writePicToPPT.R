
##' Write a picture to a ppt object.
##' 
##' @title Write a picture to a ppt object.
##' @param picfile Path of the picture to be written to ppt.
##' @param pptobject A ppt object created by \code{\link{readPPT}}.
##' @param ipage The page number.
##' @param horizontal Horizontal position to the lefttop.
##' @param vertical Vertical position to the lefttop.
##' @param height The height of the picture in ppt.
##' @param width The width of the picture in ppt.
##' @return invisible TRUE or FALSE.
##' @author Jian Li <\email{rweibo@@sina.com}>

writePicToPPT <- function(picfile, pptobject, ipage = 1, 
	horizontal = 0, vertical = 0, height = pptobject$PageSetup$SlideHeight, 
	width = pptobject$PageSetup$SlideWidth) 
{
	picfile <- normalizePath(picfile, winslash = "\\", mustWork = FALSE)
	pptslide <- pptobject[["Com"]][["Slides"]]$Item(ipage)
	
	if (pptobject$Unit == "cm") {
		horizontal <- horizontal * getOption("PIXPERCM")
		vertical <- vertical * getOption("PIXPERCM")
		height <- height * getOption("PIXPERCM")
		width <- width * getOption("PIXPERCM")
	}
	
	pptshapes <- pptslide[["Shapes"]]
	pptshapes$AddPicture(picfile, 0, -1, horizontal, vertical, width, height)
	
	invisible(TRUE)
}

