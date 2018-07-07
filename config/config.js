module.exports = {
    author: {
       label: {
           value: 'Name',
           font: {
               bold: true
           },
           fill: {
               type: 'pattern',
               pattern: 'solid',
               fgColor: { argb: 'FFC6E0B4' }
           },
           border: {
               top: { style: 'thin', color: { argb: 'FF000000' } },
               left: { style: 'thin', color: { argb: 'FF000000' } },
               bottom: { style: 'thin', color: { argb: 'FF000000' } },
               right: { style: 'thin', color: { argb: 'FF000000' } }
           }
       },
       signature: {
           font: {
               italic: true,
               underline: true
           },
           fill : {
               type: 'pattern',
               pattern: 'solid',
               fgColor: { argb: '00C6E0B4' }
           },
           value : 'Pakala,Nagateja Seetharama Srinivas Prakash',
           border: {
               top: { style: 'thin', color: { argb: 'FF000000' } },
               left: { style: 'thin', color: { argb: 'FF000000' } },
               bottom: { style: 'thin', color: { argb: 'FF000000' } },
               right: { style: 'thin', color: { argb: 'FF000000' } }
           }
       } 
    },
    header: {
        font : {
            color: { argb: 'FFFFFFFF' },
            bold: true
        },
        alignment : {
            horizontal: 'center'
        },
        fill : {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC00000' }
        },
        border: {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        }
    },
    font: {
        name: 'Calibri',
        color: { argb: 'FF000000' },
        family: 4,
        size: 11,
    },
    border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    },
    columns: ['A', 'B', 'C', 'D', 'E', 'F'],
    headervalues: ['SNO', 'Genre', 'Credit Score', 'Album Name', 'Artist', 'Release Date'],
    inputfilename: 'pakala_input.xlsx',
    outputfilename: 'pakala_Output.xlsx'
}