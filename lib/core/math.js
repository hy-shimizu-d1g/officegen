const libxslt = require('libxslt')
const fs = require('fs')
const parser = require('xml2json')
const stylesheetSource = fs.readFileSync(__dirname + '/MML2OMML.XSL', 'utf8')
console.log(__dirname)
const libmljs = libxslt.libxmljs
var stylesheetObj = libmljs.parseXml(stylesheetSource)
const stylesheet = libxslt.parse(stylesheetObj)

// tex-mathml
const { TeX } = require('mathjax-full/js/input/tex.js')
const {
  HTMLDocument
} = require('mathjax-full/js/handlers/html/HTMLDocument.js')
const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js')
const { STATE } = require('mathjax-full/js/core/MathItem.js')

const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js')
module.exports = {
  tex2mml: async function (texText) {
    const tex = new TeX({
      packages: AllPackages.filter((name) => name !== 'bussproofs')
    })

    const html = new HTMLDocument('', liteAdaptor(), { InputJax: tex })

    const {
      SerializedMmlVisitor
    } = require('mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js')
    const visitor = new SerializedMmlVisitor()
    const toMathML = (node) => visitor.visitTree(node, html)
    return toMathML(
      html.convert(texText || '', {
        display: texText,
        end: STATE.CONVERT
      })
    )
  },
  mml2omml: async function (mml) {
    var omml = await stylesheet.apply(mml)
    return omml
  },
  xml2json: function (xml) {
    return parser.toJson(xml)
  },
  _getBase: function (mathText, opts) {
    var baseMathObj = {
      'w:p': {
        '@w:raidR': '00A77427',
        '@w:rsidRDefault': '00100817',
        'm:oMathPara': {
          'm:oMathParaPr': {
            'm:jc': {
              '@m:val': 'left'
            }
          },
          'm:oMath': {
            'm:r': {
              'm:rPr': {
                'm:sty': {
                  '@m:val': 'p'
                }
              },
              'w:rPr': {
                'w:fFonts': {
                  '@w:ascii': 'Cambria Math',
                  '@w:hAnsi': 'Cambria Math'
                }
              },
              'm:t': mathText
            }
          }
        }
      }
    }
    return baseMathObj
  }
}
