# TODO: Add comment
# 
# Author: jli
###############################################################################

.onAttach <- function(libname, pkgname ){
	options(PIXPERCM = 28.35)
	tmpfolder <- system.file("tmp", package = "Rofficetool")
    rsrc <- list.files(tmpfolder, full.names = TRUE, pattern = "\\.r$", ignore.case = TRUE)
    for (i in 1:seq_along(rsrc)) {
        source(rsrc[i])
    }
	packageStartupMessage( paste("# Rofficetool Version:", utils:::packageDescription("Rofficetool", fields = "Version")) )
}

