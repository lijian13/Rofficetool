
.ppt.getText <- function(shapeobj) {
	OUT <- shapeobj[["TextFrame"]][["TextRange"]][["Text"]]
	return(OUT)
}

.ppt.getTable <- function(shapeobj) {
	nrows <- shapeobj[["Table"]][["Rows"]][["Count"]]
	ncols <- shapeobj[["Table"]][["Columns"]][["Count"]]
	OUT <- matrix("", nrow = nrows, ncol = ncols)
	for (i in 1:nrows) {
		for (j in 1:ncols) {
			OUT[i, j] <- shapeobj[["Table"]]$Cell(i, j)[["Shape"]][["TextFrame"]][["TextRange"]][["Text"]]
		}
	}
	return(OUT)
}

.ppt.getGroup <- function(shapeobj) {
	ngrpitem <- shapeobj[["GroupItems"]]$Count()
	OUT <- list()
	for (i in 1:ngrpitem) {
		OUT[[i]] <- list()
		tmp.obj <- shapeobj[["GroupItems"]]$Item(i)
		OUT[[i]][["type"]] <- tmp.obj[["Type"]]
		
		if (tmp.obj[["Type"]] %in% c(14, 17)) {
			OUT[[i]][["Value"]] <- .ppt.getText(tmp.obj)
		}
		
		if (tmp.obj[["Type"]] %in% c(19)) {
			OUT[[i]][["Value"]] <- .ppt.getTable(tmp.obj)
		}
		
		if (tmp.obj[["Type"]] %in% c(6)) {
			OUT[[i]][["Value"]] <- .ppt.getGroup(tmp.obj)
		}
	}
	return(OUT)
}
	




