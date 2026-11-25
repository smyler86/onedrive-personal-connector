import test from 'node:test';
import assert from 'node:assert/strict';
import { callDrive } from '../lib/drive.mjs';

const item = { id: 'AABB!311', name: 'image.png', size: 8, file: {mimeType:'image/png'}, eTag:'"v1"' };
const metadata = response => JSON.parse(response[0].text);

test('folder items contain durable identifiers and preserve the Microsoft pagination token', async () => {
  const args = { folder_id:'root', limit:2 };
  const graph = async path => {
    const url = new URL(path,'https://graph.microsoft.com');
    assert.equal(url.pathname,'/me/drive/root/children');
    assert.equal(url.searchParams.get('$top'),'2');
    const cursor = url.searchParams.get('$skiptoken');
    if (cursor) {
      assert.equal(cursor,'opaque+token/one');
      return Response.json({value:[]});
    }
    return Response.json({value:[item],'@odata.nextLink':'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=opaque%2Btoken%2Fone'});
  };
  const first = metadata(await callDrive('list_folder',args,graph));
  assert.equal(first.items[0].etag,item.eTag);
  assert.equal(first.items[0].id,item.id);
  const last = metadata(await callDrive('list_folder',{...args,cursor:first.next_cursor},graph));
  assert.deepEqual(last,{items:[],next_cursor:null});
});

test('photo and video previews return an image alongside original file metadata', async () => {
  const png = new Uint8Array([137,80,78,71,13,10,26,10]);
  for (const format of ['image','thumbnail']) {
    const source = format === 'thumbnail' ? {...item,name:'movie.mp4',file:{mimeType:'video/mp4'}} : item;
    const response = await callDrive('get_item',{item_id:item.id,format},async path => {
      if (path.includes('?$select=')) return Response.json(source);
      assert.equal(path,'/me/drive/items/'+item.id+(format==='thumbnail'?'/thumbnails/0/large/content':'/content'));
      return new Response(png,{headers:{'Content-Type':'image/png'}});
    });
    assert.equal(metadata(response).name,source.name);
    assert.deepEqual(response[1],{type:'image',mimeType:'image/png',data:'iVBORw0KGgo='});
  }
});

test('original binary reads use exact bytes and an oversized original stays in OneDrive', async () => {
  const file = await callDrive('get_item',{item_id:item.id,format:'file'},async path =>
    path.endsWith('/content') ? new Response(new Uint8Array([1,2,3])) : Response.json(item));
  assert.equal(file[1].resource.blob,'AQID');
  let requests = 0;
  await assert.rejects(() => callDrive('get_item',{item_id:item.id,format:'file'},async () => {
    requests++;
    return Response.json({...item,size:2*1024*1024});
  }), /download_link/);
  assert.equal(requests,1);
});

test('a folder name collision leaves the existing folder alone', async () => {
  let requests = 0;
  await assert.rejects(() => callDrive('create_folder',{parent_id:'root',name:'Photos'},async (path,init) => {
    requests++;
    assert.equal(path,'/me/drive/root/children');
    assert.equal(JSON.parse(init.body)['@microsoft.graph.conflictBehavior'],'fail');
    return new Response(null,{status:409});
  }), /409/);
  assert.equal(requests,1);
});

test('move to root resolves the true root ID and uses the displayed version', async () => {
  const result = await callDrive('update_item',{item_id:item.id,etag:item.eTag,parent_id:'root'},async (path,init) => {
    if (path.startsWith('/me/drive/root?')) return Response.json({id:'AABB!root'});
    assert.equal(new Headers(init.headers).get('If-Match'),item.eTag);
    assert.deepEqual(JSON.parse(init.body),{parentReference:{id:'AABB!root'}});
    return Response.json({...item,parentReference:{id:'AABB!root'}});
  });
  assert.equal(metadata(result).parent_id,'AABB!root');
});
