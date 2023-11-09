//参考リンク　https://qiita.com/mhgp/items/abe5679ec9e190845e60

var ExcelUtils = (function() {
  const RADIX = 26; //アルファベットの文字数
  const A = 'A'.charCodeAt(0);

  return {
    /**
     * アルファベット表記の列番号を、始まりを1とした数字にる列番号に変換する。
     * @param {string} str
     * @return {number} 
     */
    convertAlphabetColumnToNumeric: function(str) {
      var s = str.toUpperCase();
      var n = 0;
      for (var i = 0, len = s.length; i < len; i++) {
        n = (n * RADIX) + (s.charCodeAt(i) - A + 1);
      }
      return n;
    },

    /**
     * 始まりを1とした数字にる列番号を、アルファベット表記の列番号に変換する。
     * @param {number} num
     * @return {string} 
     */
    convertNumericColumnToAlphabet: function(num) {
      var n = num;
      var s = "";
      while (n >= 1) {
        n--;
        s = String.fromCharCode(A + (n % RADIX)) + s;
        n = Math.floor(n / RADIX);
      }
      return s;
    }
  };
})();

function test() {
  Logger.log(ExcelUtils.convertAlphabetColumnToNumeric("A"));   // => 1
  Logger.log(ExcelUtils.convertAlphabetColumnToNumeric("Z"));   // => 26
  Logger.log(ExcelUtils.convertAlphabetColumnToNumeric("AA"));  // => 27
  Logger.log(ExcelUtils.convertAlphabetColumnToNumeric("ZZ"));  // => 702
  Logger.log(ExcelUtils.convertAlphabetColumnToNumeric("AAA")); // => 703
  Logger.log(ExcelUtils.convertNumericColumnToAlphabet(1));      // => "A"
  Logger.log(ExcelUtils.convertNumericColumnToAlphabet(26));     // => "Z"
  Logger.log(ExcelUtils.convertNumericColumnToAlphabet(27));     // => "AA"
  Logger.log(ExcelUtils.convertNumericColumnToAlphabet(702));    // => "ZZ"
  Logger.log(ExcelUtils.convertNumericColumnToAlphabet(703));    // => "AAA"
}