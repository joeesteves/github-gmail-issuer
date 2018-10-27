function extractImagesFromBody(body: string): string[]{
  const regexp = /\!\[image\]\((https[^\!]*png)\)/g
  let link, linkAry = []
  while((link = regexp.exec(body)) !== null){
    linkAry.push(link[0].match(/https.*png/)[0])
  }
  return linkAry
}

function formatDate(date: Date) {
  return date.toISOString().substr(0,10).split('-').reverse().join('/')
}
