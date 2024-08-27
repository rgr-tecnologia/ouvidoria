import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
//import { SPHttpClient, HttpClientResponse } from "@microsoft/sp-http";
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
import "@pnp/sp/sputilities";
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
          checkbox.addEventListener('change', this.handleCheckboxChange);
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

  private handleEnviarClick = async () => {
    const tipo = (this.domElement.querySelector("#tipoInput") as HTMLSelectElement).value;
    const gravidade = (this.domElement.querySelector("#gravidadeInput") as HTMLSelectElement).value;
    const departamento = (this.domElement.querySelector("#departamentoInput") as HTMLSelectElement).value;
    const denuncia = (this.domElement.querySelector("#denunciaInput") as HTMLTextAreaElement).value;
    const anonimo = this.anonimoEstado;
    const usuario = this.usuarioLogado;

    // Verifica se os campos obrigatórios foram preenchidos
    if (!tipo || (tipo !== 'Elogio' && tipo !== 'Sugestao' && !gravidade) || !departamento || !denuncia) {
        alert('Por favor, preencha todos os campos obrigatórios.');
        return;
    }

    let destinatarios: string[];

    switch (tipo) {
        case 'Denuncia':
            if (departamento === 'Administrative' || departamento === 'Human Resources') {
                destinatarios = ['renata.alencar@cjtrade.com.br'];
            } else if (departamento === 'Legal') {
                destinatarios = ['fatima.oliveira@cjtrade.com.br'];
            } else {
                destinatarios = ['fatima.oliveira@cjtrade.com.br'];
            }
            break;
        case 'Sugestao':
            destinatarios = ['hr.br@cjtrade.net' , 'administrative.br@cjtrade.net' , 'priscila.moraes@cjtrade.com.br'];
            break;
        case 'Reclamacao':
            destinatarios = ['hr.br@cjtrade.net' , 'administrative.br@cjtrade.net' , 'priscila.moraes@cjtrade.com.br'];
            break;
        case 'Elogio':
            destinatarios = ['hr.br@cjtrade.net'];
            break;
        default:
            destinatarios = [];
            break;
    }
    await this.insertDb({
       Tipo: tipo,
       Gravidade: gravidade,
       Departamento: departamento,
       Usuario: anonimo? null :usuario,
       Anonimato: anonimo,
       Descreva: anonimo? null :denuncia,
    });
    const ccDestinatarios = [usuario];
    const emailProps: any = {
        To: destinatarios,
        Subject: `Ouvidoria - Nova ${tipo} - ${gravidade}`,
        Body: `
            <p><strong>Tipo:</strong> ${tipo}</p>
            <p><strong>Gravidade:</strong> ${gravidade}</p>
            <p><strong>Departamento:</strong> ${departamento}</p>
            <p><strong>Descrição:</strong> ${denuncia}</p>
            ${anonimo ? '' : `<p><strong>Usuário:</strong> ${usuario}</p>`}
        `,
        BCC: ccDestinatarios,
        AdditionalHeaders: {
            "content-type": "text/html"
        }
    };
    
    const sp = spfi().using(SPFx(this.context));
    await sp.utility.sendEmail(emailProps);
    

    alert('Email enviado com sucesso');
    
    location.reload();
  }

  private async insertDb(data: any) {
      const sp = spfi().using(SPFx(this.context));
      await sp.web.lists.getByTitle("Ouvidoria_New").items.add(data);
  }

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
