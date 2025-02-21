## Arquitetura do Projeto
Este projeto é uma automação baseada em Selenium para interação com páginas web. Ele possui uma arquitetura modular, separando as responsabilidades em diferentes módulos para facilitar a manutenção e escalabilidade. Seu estilo de codificação é procedural modular.
</br>
</br>
<table border="1" cellpadding="5" cellspacing="0">
  <thead>
    <tr>
      <th>Módulo</th>
      <th>Descrição</th>
      <th>Principais Responsabilidades</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td><strong>atuadorWeb</strong></td>
      <td>Módulo responsável pela interação com páginas web utilizando Selenium.</td>
      <td>
        <ul>
          <li>Controlar o navegador via Selenium.</li>
          <li>Interagir com elementos da página.</li>
          <li>Inserir arquivos.</li>
          <li>Manipular frames e alertas.</li>
          <li>Executar scripts JavaScript.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>utils</strong></td>
      <td>Módulo contendo funções auxiliares reutilizáveis para otimizar a execução.</td>
      <td>
        <ul>
          <li>Manipulação de strings.</li>
          <li>Operações com arquivos e diretórios.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>gui</strong></td>
      <td>Módulo responsável pela interface gráfica utilizando Tkinter.</td>
      <td>
        <ul>
          <li>Exibir status da execução.</li>
          <li>Disponibilizar botões de controle.</li>
          <li>Permitir seleção de arquivos e configurações pelo usuário.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>main.py</strong></td>
      <td>Script principal onde a automação é iniciada, integrando todos os módulos.</td>
      <td>
        <ul>
          <li>Orquestrar a execução da automação.</li>
          <li>Coordenar a comunicação entre os módulos.</li>
          <li>Gerenciar fluxos e exceções.</li>
        </ul>
      </td>
    </tr>
  </tbody>
</table>
</br>
</br>

### Classe Interagente:
- abrir_pagina_web(link): Abre uma página no navegador e maximiza a janela.
- interagir_pagina_web(xpath, acao, texto='', limitar_espera=False, limitar_retorno=False): Executa ações em elementos da página (cliques, preenchimentos, retornos de elementos).
- inserir_arquivo(xpath, xpath_de_espera, arquivo): Insere arquivos nos formulários da página.
- migrar_ao_frame(acao, indice=0): Alterna entre frames da página e gerencia alertas.
- interagir_javaScript(venctos_convertidos, indice, id): Interage diretamente com elementos via JavaScript.
- verificar_instabilidade(verificar): Verifica a estabilidade do carregamento dos elementos da página e força um refresh se necessário.
- fechar_driver(): Fecha o navegador ao final da automação.
  
