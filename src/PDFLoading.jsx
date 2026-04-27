import React from 'react';

export default function PDFLoading({ pdfLoading }) {
  if (!pdfLoading) return null;
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)',
      zIndex: 9999, display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center', gap: 20
    }}>
      <div style={{
        width: 56, height: 56,
        border: '6px solid rgba(255,255,255,0.3)',
        borderTop: '6px solid #fff',
        borderRadius: '50%',
        animation: 'spin 0.9s linear infinite'
      }} />
      <p style={{ color: '#fff', fontSize: 18, fontWeight: 600, margin: 0 }}>
        Generating PDF...
      </p>
      <p style={{ color: 'rgba(255,255,255,0.7)', fontSize: 13, margin: 0 }}>
        Please wait while charts are being captured
      </p>
    </div>
  );
}
