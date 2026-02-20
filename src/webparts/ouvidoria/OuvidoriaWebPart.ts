import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import styles from './OuvidoriaWebPart.module.scss';
import * as strings from 'OuvidoriaWebPartStrings';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-users/web";
import 'bootstrap/dist/css/bootstrap.min.css';

export interface IOuvidoriaWebPartProps {
  description: string;
}

export type DepartamentoItem = {
  Title: string;
  Ativo: boolean;
  Id: string;
}

export type ColaboradorItem = {
  Title: string;
  ID: string;
  Departamento: {
      Title: string;
      ID: string;
  };
}

export default class OuvidoriaWebPart extends BaseClientSideWebPart<IOuvidoriaWebPartProps> {

  _colaboradores?: ColaboradorItem[];
  _departamentos?: DepartamentoItem[];
  anonimoEstado: boolean = false;
  usuarioLogado: string = '';

  public render(): void {
    // Mapeia os valores do parâmetro 'type' para os valores do dropdown
    const typeMap: { [key: string]: string } = {
        '1': 'Denuncia',
        '2': 'Elogio',
        '3': 'Reclamacao',
        '4': 'Sugestao'
    };

    // Lê o parâmetro da URL
    const urlParams = new URLSearchParams(window.location.search);
    const typeParam = urlParams.get('type'); // 'type' é o nome do parâmetro na URL
    const tipoSelecionado = typeMap[typeParam || ''] || ''; // Mapeia o valor para o texto correspondente

    this.domElement.innerHTML = `<div class="container" style="text-align: -webkit-center;">
        <form style="width: 55%;">
            <div class="mb-3">
                <label class="form-label" style="display: flex;">Tipo *</label>
                <select class="form-select" id="tipoInput">
                    <option value="" ${!tipoSelecionado ? 'selected' : ''}>Selecione...</option>
                    <option value="Denuncia" ${tipoSelecionado === 'Denuncia' ? 'selected' : ''}>Denúncia</option>
                    <option value="Elogio" ${tipoSelecionado === 'Elogio' ? 'selected' : ''}>Elogio</option>
                    <option value="Reclamacao" ${tipoSelecionado === 'Reclamacao' ? 'selected' : ''}>Reclamação</option>
                    <option value="Sugestao" ${tipoSelecionado === 'Sugestao' ? 'selected' : ''}>Sugestão</option>
                </select> <!----> <br>
            </div>
            <div class="mb-3" id="divGravidade">
                <label class="form-label" style="display: flex;">Gravidade *</label>
                <select class="form-select" id="gravidadeInput">
                    <option value="Alta">Alta</option>
                    <option value="Media">Média</option>
                    <option value="Baixa">Baixa</option>
                </select> <!----> <br>
            </div>
            <div class="mb-3">
                <label class="form-label" style="display: flex;">Departamento *</label>
                <select class="form-select" id="departamentoInput">
                    <option value="">Selecione...</option>
                    ${this._departamentos?.map(item =>
                        `<option value="${item.Title}">
                            ${item.Title}
                        </option>`
                    ).join('')}
                </select> <!----> <br>
            </div>
            <div class="mb-3">
                <label class="form-label" style="display: flex;">Descreva *</label>
                <textarea required="required" class="form-control" id="denunciaInput"></textarea> <!----> <br>
            </div>
            <div class="mb-3">
                <!----> <input type="checkbox" id="meuCheckbox" ${this.anonimoEstado ? 'checked' : ''} required="required" style="float: left; margin-top: 0.7%;">
                <label class="form-label" style="display: flex;">&nbsp;&nbsp;Anônimo</label> <br>
            </div>
            <div class="col-12">
                <button type="button" class="btn btn-primary" id="enviarBtn">
                    Enviar
                </button>
            </div>
        </form>
    </div>`;

    this.pegaManipulacaoEventos();
    this.atualizarVisibilidadeCampos(); // Chama a função ao renderizar para garantir que a visibilidade está correta
  }

  private pegaManipulacaoEventos(): void {
      const checkbox = this.domElement.querySelector("#meuCheckbox") as HTMLInputElement;
      if (checkbox) {
          checkbox.addEventListener('change', this.handleCheckboxChange)
      }
      const enviarBtn = this.domElement.querySelector("#enviarBtn") as HTMLButtonElement;
      if (enviarBtn) {
          enviarBtn.addEventListener('click', this.handleEnviarClick);
      }
      const tipoInput = this.domElement.querySelector("#tipoInput") as HTMLSelectElement;
      if (tipoInput) {
          tipoInput.addEventListener('change', this.atualizarVisibilidadeCampos);
      }
  }

  private atualizarVisibilidadeCampos = (): void => {
      const tipo = (this.domElement.querySelector("#tipoInput") as HTMLSelectElement).value;
      const divGravidade = this.domElement.querySelector("#divGravidade") as HTMLDivElement;

      if (tipo === 'Elogio' || tipo === 'Sugestao') {
          divGravidade.style.display = 'none';
      } else {
          divGravidade.style.display = 'block';
      }
  }

  protected onInit(): Promise<void> {
      return super.onInit().then(async _ => {
          this._colaboradores = await this.carregarListaColaboradores();
          this._departamentos = await this.carregarListaDepartamentos();
          await this.obterUsuarioLogado();
      });
  }

  private handleCheckboxChange = (event: Event) => {
      const target = event.target as HTMLInputElement;
      this.anonimoEstado = target.checked;
  }


// private handleEnviarClick = async () => {
//   const tipo = (this.domElement.querySelector("#tipoInput") as HTMLSelectElement).value;
//   const gravidade = (this.domElement.querySelector("#gravidadeInput") as HTMLSelectElement).value;
//   const departamento = (this.domElement.querySelector("#departamentoInput") as HTMLSelectElement).value;
//   const denuncia = (this.domElement.querySelector("#denunciaInput") as HTMLTextAreaElement).value;
//   const anonimo = this.anonimoEstado;
//   const usuario = this.usuarioLogado;

//   // Verifica se os campos obrigatórios foram preenchidos
//   if (!tipo || (tipo !== 'Elogio' && tipo !== 'Sugestao' && !gravidade) || !departamento || !denuncia) {
//       alert('Por favor, preencha todos os campos obrigatórios.');
//       return;
//   }
  
//   // URL do Flow (gatilho HTTP)
//   const flowUrl = "https://default14393ff2969b46eeb9b8d1c3157d9e.82.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/2c5d66407c154bf4bd027cc9bcbd9c81/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=l18pcYqXZ1keqUb3CB5lSGQi2tE46zma0MzFPkeeLlw";


// try {
//   const response = await fetch(flowUrl, {
//     method: "POST",
//     headers: { "Content-Type": "application/json" },
//     body: JSON.stringify({
//       Type: tipo,
//       Gravity: gravidade,
//       Department: departamento,
//       Description: denuncia,
//       Anonimous: anonimo,
//       User: anonimo ? null : usuario
//     })
//   });

//   const text = await response.text();
//   let result: any = {};
//   try {
//     result = JSON.parse(text);
//   } catch (e) {
//     console.warn("Resposta não é JSON:", text);
//   }

//   if (response.ok && result.success) {
//     alert(result.message || "executado com sucesso");
//     location.reload();
//   } else {
//     // Usa a mensagem do JSON se existir, caso contrário usa texto bruto
//     throw new Error(result?.message || text || `Erro HTTP: ${response.status}`);
//   }

// } catch (error: any) {
//   console.error("Erro ao realizar operação:", error);
//   alert(`Erro ao realizar operação: ${error.message || error}`);
// }


// }
private handleEnviarClick = async () => {
  const enviarBtn = this.domElement.querySelector("#enviarBtn") as HTMLButtonElement;

  // Mostrar loading
  enviarBtn.disabled = true;
  const originalText = enviarBtn.textContent;
  enviarBtn.textContent = "Enviando...";

  const tipo = (this.domElement.querySelector("#tipoInput") as HTMLSelectElement).value;
  const gravidade = (this.domElement.querySelector("#gravidadeInput") as HTMLSelectElement).value;
  const departamento = (this.domElement.querySelector("#departamentoInput") as HTMLSelectElement).value;
  const denuncia = (this.domElement.querySelector("#denunciaInput") as HTMLTextAreaElement).value;
  const anonimo = this.anonimoEstado;
  const usuario = this.usuarioLogado;

  // Validação
  if (!tipo || (tipo !== 'Elogio' && tipo !== 'Sugestao' && !gravidade) || !departamento || !denuncia) {
    alert('Por favor, preencha todos os campos obrigatórios.');
    enviarBtn.disabled = false;
    enviarBtn.textContent = originalText;
    return;
  }

  const flowUrl = "https://default14393ff2969b46eeb9b8d1c3157d9e.82.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/2c5d66407c154bf4bd027cc9bcbd9c81/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=l18pcYqXZ1keqUb3CB5lSGQi2tE46zma0MzFPkeeLlw";

  try {
    const response = await fetch(flowUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json", 'x-spfx-key':'T9@vL$3xP!7kZ#2rQ^8mW&4yF*6uD157'},


      body: JSON.stringify({
        Type: tipo,
        Gravity: gravidade,
        Department: departamento,
        Description: denuncia,
        Anonimous: anonimo,
        User: anonimo ? null : usuario
      })
    });

    

    const text = await response.text();
    let result: any = {};
    try {
      result = JSON.parse(text);
    } catch (e) {
      console.warn("Resposta não é JSON:", text);
    }

    if (response.ok && result.success) {
      alert("Operação realizada com sucesso! Sua solicitação foi encaminhada aos responsáveis. Obrigado! ✅");
      location.reload();
    } else {
      throw new Error(result?.message || text || `Erro HTTP: ${response.status}`);
    }

  } catch (error: any) {
    alert("Erro! Não foi possível enviar sua solicitação. Por favor, tente novamente mais tarde. ❌");
  } finally {
    // Remove loading
    enviarBtn.disabled = false;
    enviarBtn.textContent = originalText;
  }
};



  private async obterUsuarioLogado(): Promise<void> {
      const sp = spfi().using(SPFx(this.context));
      const usuario = await sp.web.currentUser();
      this.usuarioLogado = usuario.Email;
  }

  private async carregarListaColaboradores(): Promise<ColaboradorItem[]> {
      const sp = spfi().using(SPFx(this.context));
      const response: ColaboradorItem[] = await sp.web.lists.getByTitle("Colaboradores").items.select("Title", "ID", "Departamento/Title").expand("Departamento").top(4099)();
      return response;
  }

  private async carregarListaDepartamentos(): Promise<DepartamentoItem[]> {
      const sp = spfi().using(SPFx(this.context));
      const response: DepartamentoItem[] = await sp.web.lists.getByTitle("Departamentos").items.select("Title", "Id", "Ativo").orderBy('Title').top(4099)();
      return response;
  }

  protected get dataVersion(): Version {
      return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
          pages: [
              {
                  header: {
                      description: strings.PropertyPaneDescription
                  },
                  groups: [
                      {
                          groupName: strings.BasicGroupName,
                          groupFields: [
                              PropertyPaneTextField('description', {
                                  label: strings.DescriptionFieldLabel
                              })
                          ]
                      }
                  ]
              }
          ]
      };
  }
}






