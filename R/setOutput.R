
##' Initialize the output file.
##' 
##' @title Initialize the output file by copying the template file.
##' @param templatefile Path of the template file.
##' @param outfile Path of the output file.
##' @return A string of the path.
##' @author Jian Li <\email{rweibo@@sina.com}>


setOutput <- function(templatefile, outfile = NULL) 
{
	if (is.null(outfile)) {
		f1 <- gsub("\\.[^\\.]*$", "", basename(templatefile))
		f2 <- gsub("^.*\\.", "", basename(templatefile))
		outfile <- file.path(getwd(), paste0(f1, "_", format(Sys.time(), "%Y%m%d%H%M%S"), ".", f2))
		outfile <- normalizePath(outfile, winslash = "/", mustWork = FALSE)
	} else {
		outfile <- normalizePath(outfile, winslash = "/", mustWork = FALSE)
		if (!file.exists(dirname(outfile))) {
			dir.create(dirname(outfile), showWarnings = FALSE, recursive = TRUE)
		}
	}
    
    file.copy(from = templatefile, to = outfile, overwrite = TRUE)
    return(outfile)
}

