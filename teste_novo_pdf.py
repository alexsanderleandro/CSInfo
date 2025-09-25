#!/usr/bin/env python3
# Teste da nova funcionalidade PDF reorganizada

from csinfo import write_pdf_report

# Dados de teste mais completos
test_lines = [
    "Relatório gerado: 2025-09-25T15:30:00.000000",
    "Nome do computador: CEOSOFT-059",
    "Tipo: Desktop",
    "Versão do sistema operacional: Microsoft Windows 11 Pro (Version 10.0.26100) - 64 bits",
    "Ativação do Windows: Ativado",
    "Rede: Domínio: ceosoftwaread.net",
    "Processador 1: Intel(R) Core(TM) i7-7700 CPU @ 3.60GHz",
    "  Fabricante: GenuineIntel | Arquitetura: x64",
    "Memória RAM total: 7.87 GB",
    "Pente de Memória 1: 8 GB | DDR4 | 2400 MHz | DIMM",
    "Disco 1: ST1000DM010-2EP102 | Tamanho: 931.51 GB",
    "Unidade C: (Sem rótulo) | Total: 893.51 GB | Usado: 539.82 GB | Livre: 353.69 GB | Sistema: NTFS",
    "Versão do Office: Microsoft Office Professional Plus 2019 - pt-br 16.0.19127.20240",
    "Ativação do Office: Ativado",
    "Placa mãe: ASUSTeK COMPUTER INC. | Modelo: H110M-C/BR | Serial: 180414466200454",
    "Monitor 1: GSM | Modelo: LG FULL HD | Serial: 16843009",
    "Placa de Rede 1: Realtek PCIe GbE Family Controller | Fabricante: Realtek | Velocidade: 100 Mbps",
    "Placa de Vídeo 1: NVIDIA GeForce GT 220 | Fabricante: NVIDIA | Memória: 1 GB | Tipo: Offboard",
    "Teclado conectado: SIM",
    "Mouse conectado: SIM",
    "SQL Server 1: Instância: Default | Versão: 16.0.1000.6 | Status: Stopped",
    "Antivírus 1: Kaspersky | Status: Ativado",
    "Administrador 1: Administrador",
    "Administrador 2: alex",
    "Administrador 3: CEOSOFTWARE",
    "Impressora 1: HP LaserJet 1020 | Serial/ID: NÃO OBTIDO",
    "",
    "=== SOFTWARES INSTALADOS ===",
    "  1. Android Studio | Versão: 2025.1 | Editor: Google LLC",
    "  2. Google Chrome | Versão: 140.0.7339.187 | Editor: Google LLC",
    "  3. Microsoft Office Professional Plus 2019 - pt-br | Versão: 16.0.19127.20240 | Editor: Microsoft Corporation",
    "  4. WinRAR 7.13 (64-bit) | Versão: 7.13.0 | Editor: win.rar GmbH"
]

print("Testando nova geração de PDF organizada...")
result = write_pdf_report("teste_novo_pdf.pdf", test_lines, "CEOSOFT-059")

if result:
    print("✅ PDF reorganizado gerado com sucesso!")
    print("Arquivo: teste_novo_pdf.pdf")
else:
    print("❌ Erro ao gerar PDF reorganizado")