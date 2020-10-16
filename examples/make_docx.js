var async = require('async')
var officegen = require('../')

var fs = require('fs')
var path = require('path')
const { result } = require('lodash')
const { doc } = require('prettier')

var outDir = path.join(__dirname, '../tmp/')

// var themeXml = fs.readFileSync(path.resolve(__dirname, 'themes/testTheme.xml'), 'utf8')

var docx = officegen({
  type: 'docx',
  orientation: 'portrait',
  pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
  // The theme support is NOT working yet...
  // themeXml: themeXml
})

// Remove this comment in case of debugging Officegen:
// officegen.setVerboseMode ( true )

docx.on('error', function (err) {
  console.log(err)
})
var pObj = docx.createP()

pObj.addText('Simple')
pObj.addText(' with color', { color: '000088' })
pObj.addText(' and back color.', { color: '00ffff', back: '000088' })

pObj = docx.createP()

pObj.addText('Since ')
pObj.addText('officegen 0.2.12', {
  back: '00ffff',
  shdType: 'pct12',
  shdColor: 'ff0000'
}) // Use pattern in the background.
pObj.addText(' you can do ')
pObj.addText('more cool ', { highlight: true }) // Highlight!
pObj.addText('stuff!', { highlight: 'darkGreen' }) // Different highlight color.

pObj = docx.createP()

pObj.addText('Even add ')
pObj.addText('external link', { link: 'https://github.com' })
pObj.addText('!')

pObj = docx.createP()

pObj.addText('Bold + underline', { bold: true, underline: true })

pObj = docx.createP({ align: 'center' })

pObj.addText('Center this text', {
  border: 'dotted',
  borderSize: 12,
  borderColor: '88CCFF'
})

pObj = docx.createP()
pObj.options.align = 'right'

pObj.addText('Align this text to the right.')

pObj = docx.createP()

pObj.addText(
  'Those two lines are in the same paragraph,\nbut they are separated by a line break.'
)

docx.putPageBreak()

pObj = docx.createP()

pObj.addText('Fonts face only.', { font_face: 'Arial' })
pObj.addText(' Fonts face and size.', { font_face: 'Arial', font_size: 40 })

pObj = docx.createP()

pObj.addText('בדיקה האם אפשר לכתוב טקסט בעברית', { rtl: true })

docx.putPageBreak()

pObj = docx.createP()

pObj.addImage(path.resolve(__dirname, 'images_for_examples/image3.png'))

docx.putPageBreak()

pObj = docx.createP()

pObj.addImage(path.resolve(__dirname, 'images_for_examples/image1.png'))

pObj = docx.createP()

pObj.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'))
pObj.addImage(path.resolve(__dirname, 'images_for_examples/sword_002.png'))
pObj.addImage(path.resolve(__dirname, 'images_for_examples/sword_003.png'))
pObj.addText('... some text here ...', { font_face: 'Arial' })
pObj.addImage(path.resolve(__dirname, 'images_for_examples/sword_004.png'))

pObj = docx.createP()

pObj.addImage(path.resolve(__dirname, 'images_for_examples/image1.png'))

docx.putPageBreak()

// Add an unordered list
pObj = docx.createListOfDots()

pObj.addText('Unordered Option 1')

pObj = docx.createNestedUnOrderedList({
  level: 2
})

pObj.addText('Unordered Nested Option 1')

pObj = docx.createListOfDots()

pObj.addText('Unordered Option 2')

docx.putPageBreak()

// Add an ordered list
pObj = docx.createListOfNumbers()

pObj.addText('Ordered Option 1')

// Add a nested list
pObj = docx.createNestedOrderedList({
  level: 2
})

pObj.addText('Second Level Option 1')

// Add a nested third level list
pObj = docx.createNestedOrderedList({
  level: 3
})

pObj.addText('Third Level Option 1')

pObj = docx.createNestedOrderedList({
  level: 2
})

pObj.addText('Second Level Option 2')

pObj = docx.createListOfNumbers()

pObj.addText('Ordered Option 2')

pObj.addHorizontalLine()

pObj = docx.createP({ backline: 'E0E0E0' })

pObj.addText('Backline text1')

pObj.addText(' text2')

pObj = docx.createP()

pObj.addText('Strikethrough text', { strikethrough: true })

pObj.addText('superscript', { superscript: true })
pObj.addText('subscript', { subscript: true })

var table = [
  [
    {
      val: 'No.',
      opts: {
        cellColWidth: 4261,
        b: true,
        sz: '48',
        shd: {
          fill: '7F7F7F',
          themeFill: 'text1',
          themeFillTint: '80'
        },
        fontFamily: 'Avenir Book'
      }
    },
    {
      val: 'Title1',
      opts: {
        b: true,
        color: 'A00000',
        align: 'right',
        shd: {
          fill: '92CDDC',
          themeFill: 'text1',
          themeFillTint: '80'
        }
      }
    },
    {
      val: 'Title2',
      opts: {
        align: 'center',
        cellColWidth: 42,
        b: true,
        sz: '48',
        shd: {
          fill: '92CDDC',
          themeFill: 'text1',
          themeFillTint: '80'
        }
      }
    }
  ],
  [1, 'All grown-ups were once children', ''],
  [2, 'there is no harm in putting off a piece of work until another day.', ''],
  [
    3,
    'But when it is a matter of baobabs, that always means a catastrophe.',
    ''
  ],
  [4, 'watch out for the baobabs!', 'END']
]

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: 'ada',
  tableAlign: 'left',
  tableFontFamily: 'Comic Sans MS'
}

pObj = docx.createTable(table, tableStyle)

var mathSample = [
  'y = x * 2',
  '\\frac{\\pi}{2} = \\left( \\int_{0}^{\\infty} \\frac{\\sin x}{\\sqrt{x}} dx \\right)^2 =\\sum_{k=0}^{\\infty} \\frac{(2k)!}{2^{2k}(k!)^2} \\frac{1}{2k+1} =\\prod_{k=1}^{\\infty} \\frac{4k^2}{4k^2 - 1}'
]

async.series(
  [
    async function (done) {
      var mml = await docx.tex2mml(mathSample[1])
      // console.log(mml)
      var omml = await docx.mml2omml(mml)
      // console.log(omml)
      return omml
    }
  ],
  (err, result) => {
    result.forEach((value, index) => {
      docx.createMath(value)
    })
    var out = fs.createWriteStream(path.join(outDir, 'example.docx'))
    async.parallel(
      [
        function (done) {
          out.on('close', function () {
            console.log('Finish to create a DOCX file.')
            done(null)
          })
          docx.generate(out)
        }
      ],
      function (err) {
        if (err) {
          console.log('error: ' + err)
        } // Endif.
      }
    )
  }
  // function (result) {
  //   console.log(result)
  //   docx.createMath(result, {})

  //   out.on('error', function (err) {
  //     console.log(err)
  //   })

  //   async.parallel(
  //     [
  //       function (done) {
  //         out.on('close', function () {
  //           console.log('Finish to create a DOCX file.')
  //           done(null)
  //         })
  //         docx.generate(out)
  //       }
  //     ],
  //     function (err) {
  //       if (err) {
  //         console.log('error: ' + err)
  //       } // Endif.
  //     }
  //   )
  // }
)
// pObj = docx.createMath()
