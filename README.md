# Automação Documentos
## Abonada
Durante meu estagio, notei que faziamos muitos documentos repetidos. Em princípio, usei o documento de abono para automatizar, e que serviria de modelo para alguns outros documentos.

A proposta era trocar as informações necessarias no modelo base em áreas pré definidas.
- Matricula
- Nome
- Cargo
- Data do Abono
- Data do Documento

E precisaria apenas das seguintes informações para a criação do documento:
- Matricula ou nome
- data que gostaria de abonar.

Para então, verificar se as condições do pedido estão validas e por fim finalizar o documento.

Utilizava-se uma tabela csv com as informações das pessoas do local e trocar as informações.

## Para criação da Abonada utilizei:
- Python (para validação e automatização)
- Arquivo CSV (como Banco de dados)
- Arquivo DOCX (modelo base e documento final)

