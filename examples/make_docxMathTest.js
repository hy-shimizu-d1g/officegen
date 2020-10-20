var async = require('async')
var officegen = require('../')

var fs = require('fs')
var path = require('path')
const { doc } = require('prettier')
const { column } = require('../lib/pptx/charts')

var outDir = path.join(__dirname, '../tmp/')

// var themeXml = fs.readFileSync(path.resolve(__dirname, 'themes/testTheme.xml'), 'utf8')

var docx = officegen({
  type: 'docx',
  orientation: 'portrait',
  author: 'hy-shimizu',
  pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 },
  columns: 2
  // The theme support is NOT working yet...
  // themeXml: themeXml
})

// Remove this comment in case of debugging Officegen:
// officegen.setVerboseMode ( true )

docx.on('error', function (err) {
  console.log(err)
})

var docTextData = {
  title: 'this is Math Test Docx',
  contents: [
    'first Contents \nsecond line',
    '\\frac{\\pi}{2} = \\left( \\int_{0}^{\\infty} \\frac{\\sin x}{\\sqrt{x}} dx \\right)^2 =\\sum_{k=0}^{\\infty} \\frac{(2k)!}{2^{2k}(k!)^2} \\frac{1}{2k+1} =\\prod_{k=1}^{\\infty} \\frac{4k^2}{4k^2 - 1}',
    'inline math $\\frac{\\pi}{2} = \\left( \\int_{0}^{\\infty} \\frac{\\sin x}{\\sqrt{x}} dx \\right)^2 =\\sum_{k=0}^{\\infty} \\frac{(2k)!}{2^{2k}(k!)^2} \\frac{1}{2k+1} =\\prod_{k=1}^{\\infty} \\frac{4k^2}{4k^2 - 1}$ テキスト',
    '日本語も入力は可能$\\frac{\\pi}{2k+1}$'
  ],
  contents2: 'test'
}
var makeOmml = async function (tex) {
  var mml = await docx.tex2mml(tex)
  var omml = await docx.mml2omml(mml)
  return omml
}

async.series(
  [
    async function () {
      var pObj = docx.createP()
      await pObj.addText(docTextData.title)
    },
    async function () {
      var con = docTextData.contents
      for (var i = 0; i < con.length; i++) {
        var pObj = docx.createP()
        var val = con[i]
        var index = i
        if (val.match(/^\\/)) {
          var omml = await makeOmml(val)
          pObj.addMath(omml.replace(/\<\?.*?\n/, ''))
        } else if (val.match(/\$\\/)) {
          var split = val.split(/\$/)
          for (var t = 0; t < split.length; t++) {
            if (split[t].match(/^\\/)) {
              var inlineOmml = await makeOmml(split[t])
              pObj.addMath(inlineOmml.replace(/\<\?.*?\n/, ''))
            } else {
              if (split[t] !== '') {
                pObj.addText(split[t])
              }
            }
          }
        } else {
          pObj.addText(val)
        }
      }
      // await con.map(async (val, index) => {
      //   console.log('exec => ' + index)
      //   if (val.match(/^\\/)) {
      //     var omml = await makeOmml(val)
      //     await docx.createObject(omml)
      //     console.log(index)
      //     // } else if (docTextData[key].match(/\$\\/)) {
      //     //   var split = docTextData[key].split(/\$/);
      //     //   for (var i = 0; i < split.length; i++) {
      //     //     if (split[i].mathch(/^\\/)) {
      //     //       var omml = await makeOmml(split[i])
      //     //       docx.createObject(omml)
      //     //     } else {
      //     //       pObj.addText(split[i]);
      //     //     }
      //     //   }
      //   } else {
      //     await pObj.addText(val)
      //     console.log(index)
      //   }
      // })
      // )
    }
  ],
  (err) => {
    console.log(err)
    var out = fs.createWriteStream(path.join(outDir, 'sample.docx'))
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
)

// async.series(
//   [
//     async function (done) {
//       var mml = await docx.tex2mml(mathSample[1])
//       // console.log(mml)
//       var omml = await docx.mml2omml(mml)
//       // console.log(omml)
//       return omml
//     }
//   ],
//   (result) => {
//     result.forEach((value, index) => {
//       docx.createObject(value)
//     })
//     var out = fs.createWriteStream(path.join(outDir, 'example.docx'))
//     async.parallel(
//       [
//         function (done) {
//           out.on('close', function () {
//             console.log('Finish to create a DOCX file.')
//             done(null)
//           })
//           docx.generate(out)
//         }
//       ],
//       function (err) {
//         if (err) {
//           console.log('error: ' + err)
//         } // Endif.
//       }
//     )
//   }
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
// )
// pObj = docx.createMath()
