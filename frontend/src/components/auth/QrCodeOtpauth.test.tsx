import { describe, expect, it } from 'vitest'
import { render, screen } from '@testing-library/react'
import { QrCodeOtpauth } from '@/components/auth/QrCodeOtpauth'

describe('QrCodeOtpauth', () => {
  it('gera o QR code como imagem SVG', async () => {
    render(<QrCodeOtpauth uri="otpauth://totp/Revalle:x?secret=JBSWY3DPEHPK3PXP" />)

    const img = await screen.findByAltText('QR code para o aplicativo autenticador')
    expect(img.getAttribute('src')).toMatch(/^data:image\/svg\+xml/)
  })
})
