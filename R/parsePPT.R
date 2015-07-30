
##' Get details of a PPT file.
##' 
##' @title Get details of a PPT object.
##' @param pptobject A ppt object created by \code{\link{readPPT}}.
##' @return A list.
##' @author Jian Li <\email{rweibo@@sina.com}>

parsePPT <- function(pptobject)
{
	if (!inherits(pptobject, "PPTobj")) {
		stop("Wrong PPT object!\nPlease use 'readPPT' to create it.")
	}
	
	if (pptobject$Unit == "cm") {
		unit.f <- getOption("PIXPERCM")
	} else {
		unit.f <-  1
	}
	
	SHAPETYPE <- c("AutoShape", "Callout", "Chart", "Comment",
			"Freeform", "Group", "EmbeddedOLEObject", "FormControl", 
			"Line", "LinkedOLEObject", "LinkedPicture", "OLEControlObject", 
			"Picture", "Placeholder", "TextEffect", "Media",
			"TextBox", "ScriptAnchor", "Table", "Canvas", 
			"N/A", "N/A", "N/A", "Diagram")
	
	pptout <- list()
	for (i in 1:pptobject$PageNum) {
		pptslide <- pptobject[["Com"]][["Slides"]]$Item(i)
		nshapes <- pptslide[["Shapes"]][["Count"]]
		shapesnm <- character()
		for (j in 1:nshapes) {
			shapesnm[j] <- paste("[", j, "]. ", SHAPETYPE[pptslide[["Shapes"]][[j]][["Type"]]],
					": ", pptslide[["Shapes"]][[j]][["Name"]],
					sep = "")
		}
		pptout[[i]] <- lapply(1:nshapes, FUN = function(X) list())
		names(pptout[[i]]) <- shapesnm
		
		for (j in 1:nshapes) {
			pptshape <- pptslide[["Shapes"]]$Item(j)
			pptout[[i]][[j]][["Type"]] <- pptshape[["Type"]]		
			pptout[[i]][[j]][["Position"]] <- c(
					Left = pptshape[["Left"]] / unit.f,
					Top = pptshape[["Top"]] / unit.f,
					Width = pptshape[["Width"]] / unit.f,
					Height = pptshape[["Height"]] / unit.f
			)
			if (pptshape[["Type"]] %in% c(14, 17)) {
				pptout[[i]][[j]][["Value"]] <- .ppt.getText(pptshape)
			}
			
			if (pptshape[["Type"]] %in% c(19)) {
				pptout[[i]][[j]][["Value"]] <- .ppt.getTable(pptshape)
			}
			
			if (pptshape[["Type"]] %in% c(6)) {
				pptout[[i]][[j]][["Value"]] <- .ppt.getGroup(pptshape)
			}
		}
	}
	names(pptout) <- paste0("page", 1:length(pptout))
	return(pptout)
}

