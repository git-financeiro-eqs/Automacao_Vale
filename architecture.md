## Arquitetura do Projeto
Este projeto é uma automação baseada em Selenium para interação com páginas web. Ele possui uma arquitetura modular, separando as responsabilidades em diferentes módulos para facilitar a manutenção e escalabilidade. Seu estilo de codificação é procedural modular.
</br>
</br>
<table>
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
      <td>Módulo responsável pela interação com páginas web utilizando Selenium, através da classe Interagente.</td>
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
          <li>Disponibilizar botões de controle.</li>
          <li>Permitir seleção de arquivos e configurações pelo usuário.</li>
        </ul>
      </td>
    </tr>
    <tr>
      <td><strong>main.py</strong></td>
      <td>Script principal onde a função da automação está, integrando os módulos utils e atuadorWeb.</td>
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
- abrir_pagina_web(): Abre uma página no navegador e maximiza a janela.
- interagir_pagina_web(): Executa ações em elementos da página (cliques, preenchimentos, retornos de elementos).
- inserir_arquivo(): Insere arquivos nos formulários da página.
- migrar_ao_frame(): Alterna entre frames da página e gerencia alertas.
- interagir_javaScript(): Interage diretamente com elementos via JavaScript.
- verificar_instabilidade(): Verifica a estabilidade do carregamento de alguns elementos específicos da página e força um refresh se necessário.
- fechar_driver(): Fecha o navegador ao final da automação.
</br>
</br>

### Observação:

Foi necessário uma engenharia lógica para superar a instabilidade do servidor da Vale. A todo momento os elementos da pagina "piscam", saem da tela e voltam, e como essa automação é um script procedural, corria o risco constante de no momento de execução de um comando do script o elemento manipulado da página sair da tela, fazendo com que a sequencia algorítmica se perdesse. A solução encontrada foi ativar uma função paralela, uma thread, que é executada concomitantemente à função de preencher o formulário da pagina. Essa função é a verificar_portal(), presente no módulo main.py. Ela se utiliza de uma lista queue chamada controle, e sua lógica de funcionamento é bem simples: ela está em um loop infinito de verificação, constantemente verificando se um elemento referencial está presente na pagina. Se acaso o elemento surge, ela insere um valor na lista controle. A lista controle é monitorada em momentos críticos de execução da função principal e, caso ela não esteja mais vazia, a pagina é reiniciada com o comando de teclado "f5", em seguida, a lista controle é zerada e a função alimentar_portal_vale(), que é uma função recursiva, retorna ela mesma reiniciando o processo para aquele faturamento que está sendo inserido no momento.
