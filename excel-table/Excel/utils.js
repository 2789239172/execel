// 添加字母映射
const mapLetter = {}
for (var i = 0; i < 26; i++) {
  mapLetter[String.fromCharCode((65 + i))] = i + 1
}

// 数字转字母
export const numToLetter = (num) => {
  let arr = [];
  if (num === 0) return "A";
  while (num) {
    let letter = String.fromCharCode(65 + Math.floor(num % 26));
    arr.push(letter);
    num = Math.floor(num / 26);
  }
  return arr.reverse().join("");
}

// 获取合并位置信息
export const getRange = (rangeStr) => {
  const data = rangeStr.replace(':', '').split('$').slice(1)
  const x1 = data[0]
  const x2 = data[2]
  const y1 = data[1]
  const y2 = data[3]

  //  进制转换
  const getIndex = label => {
    return label.split('')
              .reverse()
              .reduce((total, item, i) => total + mapLetter[item] * 26 ** i, 0) - 1
  }

  return [
    [getIndex(x1), getIndex(x2)],
    [y1 - 1, y2 - 1]
  ]
}

// 导出
export const downloadExcel = (buffer, name = 'excel') => {
  const ba = new Blob([buffer])
  const a = document.createElement('a')
  a.href = URL.createObjectURL(ba) // 创建对象超链接
  a.download = name + new Date().getTime() + '.xlsx'
  a.target = '_block'
  document.body.appendChild(a)
  a.click() // 模拟点击实现下载
  document.body.removeChild(a)
  setTimeout(function() { // 延时释放
      URL.revokeObjectURL(ba) // 用URL.revokeObjectURL()来释放这个object URL
  }, 100)
}

export default {
  numToLetter,
  getRange,
  downloadExcel
}

// //  start
// const rangeStr = "$A$5:$B$9"
// console.log(getRange(mapLetter, rangeStr))