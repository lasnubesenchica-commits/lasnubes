function debugEmail() {
  const threads = GmailApp.search(
    'from:automated@airbnb.com subject:"Reserva confirmada:" newer_than:90d'
  );
  // Toma el primer email que falla (ej. Aneea Toomajian)
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const body = msg.getPlainBody();
      if (body.includes('TOOMAJIAN') || body.includes('Toomajian')) {
        Logger.log('=== BODY COMPLETO ===');
        Logger.log(body.substring(0, 3000));
      }
    });
  });
}