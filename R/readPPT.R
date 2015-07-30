
##' Read a PPT file into COM object.
##' 
##' @title Read a PPT file into COM object.
##' @param pptfile Path of the ppt.
##' @param unit Unit of the length.
##' @return A 'PPTobj' object.
##' @author Jian Li <\email{rweibo@@sina.com}>

readPPT <- function(pptfile, unit = c("cm", "px"))
{
	unit <- match.arg(unit)
	if (unit == "cm" && !is.null(getOption("PIXPERCM"))) {
		unit.f <- getOption("PIXPERCM")
	} else {
		unit <- "px"
		unit.f <- 1
	}
	
	pptout <- list()
	pptfile <- normalizePath(pptfile, winslash = "/", mustWork = TRUE)
	pptapp <- COMCreate("PowerPoint.Application")
	pptpre <- pptapp[["Presentations"]]$Open(pptfile)
	
	pptout[["Name"]] <- pptpre[["FullName"]]
	pptout[["Unit"]] <- unit
	pptout[["PageSetup"]] <- list(
			SlideOrientation = c("Landscape", "Portrait")[pptpre[["PageSetup"]][["SlideOrientation"]]],
			SlideSize = pptpre[["PageSetup"]][["SlideSize"]],
			SlideHeight = pptpre[["PageSetup"]][["SlideHeight"]] / unit.f,
			SlideWidth = pptpre[["PageSetup"]][["SlideWidth"]] / unit.f
	)
	pptout[["PageNum"]] <- pptpre[["Slides"]][["Count"]]
	pptout[["Com"]] <- pptpre
	
	class(pptout) <- c("PPTobj", "Rofficetool")
	return(pptout)
	
}

