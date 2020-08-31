
    git checkout master -- bits/47_styxml.js

    sed -i "s|xlscfb.flow.js|dist/xlscfb.js|g" Makefile

    echo bits/93_node.js >> mini.lst
    echo bits/94_xmlbuilder.js >> mini.lst

    sed -i -e 's/95_api/92_api/g' mini.lst
    echo bits/95_stylebuilder.js >> mini.lst
    cd bits
    mv 95_api.js 92_api.js
    mv 97_node.js 93_node.js
    mv 91_xmlbuilder.js 94_xmlbuilder.js 
    mv 92_stylebuilder.js 95_stylebuilder.js



