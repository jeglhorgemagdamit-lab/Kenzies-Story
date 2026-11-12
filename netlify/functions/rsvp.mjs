// Private same-origin bridge: browser -> Netlify Function -> Google Apps Script -> private Sheet.
export default async function handler(request) {
  if(request.method !== 'POST') return Response.json({ok:false,message:'Method not allowed.'},{status:405});
  const origin = request.headers.get('origin');
  const siteOrigin = new URL(request.url).origin;
  if(origin && origin!==siteOrigin) return Response.json({ok:false,message:'Invalid origin.'},{status:403});
  const endpoint = process.env.GOOGLE_APPS_SCRIPT_URL;
  const secret = process.env.RSVP_SHARED_SECRET;
  if(!endpoint || !secret) return Response.json({ok:false,message:'RSVP service is not configured yet.'},{status:503});
  let data;
  try { if(Number(request.headers.get('content-length')||0)>12000)throw Error('Too large'); data=await request.json(); }
  catch { return Response.json({ok:false,message:'Invalid request.'},{status:400}); }
  if(!data || !['lookup','submit'].includes(data.action)) return Response.json({ok:false,message:'Invalid action.'},{status:400});
  if(typeof data.code !== 'string' || data.code.length>64) return Response.json({ok:false,message:'Invalid invitation code.'},{status:400});
  if(data.action==='submit' && (!Array.isArray(data.guests)||data.guests.length>35))return Response.json({ok:false,message:'Invalid guest list.'},{status:400});
  try {
    const controller=new AbortController();
    const timer=setTimeout(()=>controller.abort(),22000);
    let response;
    try{
      response=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({...data,secret}),redirect:'follow',signal:controller.signal});
    }finally{clearTimeout(timer);}
    if(!response.ok)return Response.json({ok:false,message:'Google Sheets is temporarily unavailable.'},{status:502});
    const result=await response.json();
    return Response.json(result,{status:result.ok?200:400,headers:{'Cache-Control':'no-store'}});
  }catch(err){console.error('RSVP proxy error:',err);return Response.json({ok:false,message:'Could not reach the guest list. Please try again.'},{status:502});}
}
export const config = {path:'/api/rsvp'};
