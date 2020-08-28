    echo "hello"
    
    #adding bits file names to mini.lst
    echo bits/93_node.js >> mini.lst
    echo bits/94_xmlbuilder.js >> mini.lst
    #replace the existing file name where needed
    sed -i -e 's/95_api/92_api/g' mini.lst
    echo bits/95_stylebuilder.js >> mini.lst
    #renaming existng files to maintain the sequence of code being generated in final xlsx file
    cd bits
    mv 95_api.js 92_api.js
    mv 97_node.js 93_node.js
    mv 91_xmlbuilder.js 94_xmlbuilder.js 
    mv 92_stylebuilder.js 95_stylebuilder.js