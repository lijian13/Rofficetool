# TODO: Add comment
# 
# Author: jli
###############################################################################

.onAttach <- function(libname, pkgname ){
	options(PIXPERCM = 28.35)
	packageStartupMessage( paste("# Rofficetool Version:", utils:::packageDescription("Rofficetool", fields = "Version")) )
}

