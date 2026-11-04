function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    
    if (params.action === 'saveLead') {
      return guardarLead(params.lead);
    }
    
    return jsonResponse({ error: 'Acción no válida' });
  } catch (error) {
    return jsonResponse({ error: error.message });
  }
}
