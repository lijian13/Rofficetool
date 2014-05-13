
##' Get details of a PPT file.
##' 
##' @title Get details of a PPT file.
##' @param pptfile Path of the ppt.
##' @return A list.
##' @author Jian Li <\email{rweibo@@sina.com}>

# library("RDCOMClient")
# pptfile = "E:\\Mango\\Training\\Youku\\dev\\test\\ppt3.pptx"
parsePPT <- function(pptfile)
{
	pptout <- list()
	pptfile <- normalizePath(pptfile, winslash = "/", mustWork = TRUE)
	pptapp <- COMCreate("PowerPoint.Application")
	pptpre <- pptapp[["Presentations"]]$Open(pptfile)
	
	pptout[["Name"]] <- pptpre[["FullName"]] 
	pptout[["PageSetup"]] <- list(
			SlideOrientation = c("Landscape", "Portrait")[pptpre[["PageSetup"]][["SlideOrientation"]]],
			SlideSize = pptpre[["PageSetup"]][["SlideSize"]],
			SlideHeight = pptpre[["PageSetup"]][["SlideHeight"]] / getOption("PIXPERCM"),
			SlideWidth = pptpre[["PageSetup"]][["SlideWidth"]] / getOption("PIXPERCM")
	)
	pptout[["Number of slides"]] <- pptpre[["Slides"]][["Count"]]
	pptout[["Slides"]] <- list()
	
	SHAPETYPE <- c("AutoShape", "Callout", "Chart", "Comment",
			"Freeform", "Group", "EmbeddedOLEObject", "FormControl", 
			"Line", "LinkedOLEObject", "LinkedPicture", "OLEControlObject", 
			"Picture", "Placeholder", "TextEffect", "Media",
			"TextBox", "ScriptAnchor", "Table", "Canvas", 
			"N/A", "N/A", "N/A", "Diagram")
	
	for (i in 1:pptout[["Number of slides"]]) {
		pptslide <- pptpre[["Slides"]]$Item(i)
		nshapes <- pptslide[["Shapes"]][["Count"]]
		shapesnm <- paste(SHAPETYPE[sapply(pptslide[["Shapes"]], "[[", "Type")],
				": ", sapply(pptslide[["Shapes"]], "[[", "Name"),
				sep = "")
		pptout[["Slides"]][[i]] <- lapply(1:nshapes, FUN = function(X) list())
		names(pptout[["Slides"]][[i]]) <- shapesnm
		
		for (j in 1:nshapes) {
			pptshape <- pptslide[["Shapes"]]$Item(j)
			pptout[["Slides"]][[i]][[j]][["Type"]] <- pptshape[["Type"]]		
			pptout[["Slides"]][[i]][[j]][["Position"]] <- c(
					Left = pptshape[["Left"]] / getOption("PIXPERCM"),
					Top = pptshape[["Top"]] / getOption("PIXPERCM"),
					Width = pptshape[["Width"]] / getOption("PIXPERCM"),
					Height = pptshape[["Height"]] / getOption("PIXPERCM")
			)
			if (pptshape[["Type"]] %in% c(14, 17)) {
				pptout[["Slides"]][[i]][[j]][["Value"]] <- .ppt.getText(pptshape)
			}
			
			if (pptshape[["Type"]] %in% c(19)) {
				pptout[["Slides"]][[i]][[j]][["Value"]] <- .ppt.getTable(pptshape)
			}
			
			if (pptshape[["Type"]] %in% c(6)) {
				pptout[["Slides"]][[i]][[j]][["Value"]] <- .ppt.getGroup(pptshape)
			}
		}
	}
	
	pptpre$Close()
	pptapp$Quit()
	return(pptout)
	
}

