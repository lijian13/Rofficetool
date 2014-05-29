
##' Get details of a PPT file.
##' 
##' @title Get details of a PPT file.
##' @param pptfile Path of the ppt.
##' @return A list.
##' @author Jian Li <\email{rweibo@@sina.com}>


copyFromTemplate <- function(templatefile, outfile) 
{
    outfile <- normalizePath(outfile, winslash = "/", mustWork = FALSE)
    file.copy(from = templatefile, to = outfile, overwrite = TRUE)
    return(outfile)
}

