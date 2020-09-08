    rem 47_styxml has a function 'write_numFmts' which causes styles to break when any changes are made to it, hence we always preserve the master version
    git checkout master -- bits/47_styxml.js

    rem the file path in sheetjs is not applicable to this fork hence correcting the file path in makefile as well
    sed -i "s|xlscfb.flow.js|dist/xlscfb.js|g" Makefile

    rem mini.lst needs all the bits file names mentioned hence correcting the file names already present and adding ones not listed
    rem these commands must always be executed before the renaming ones
    echo bits/93_node.js >> mini.lst
    echo bits/94_xmlbuilder.js >> mini.lst
    sed -i -e 's/95_api/92_api/g' mini.lst
    echo bits/95_stylebuilder.js >> mini.lst

    rem some files need renaming so that when we run 'make' and generate xlsx.js the sequence of functions is desired one. The desired sequence must have 'var utils' right above 'function(utils)' 
    cd bits
    mv 95_api.js 92_api.js
    mv 97_node.js 93_node.js
    mv 91_xmlbuilder.js 94_xmlbuilder.js 
    mv 92_stylebuilder.js 95_stylebuilder.js







