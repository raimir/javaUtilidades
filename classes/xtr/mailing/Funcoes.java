package xtr.mailing;

import lotus.domino.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Vector;
import org.json.*;

//import xtr.main.Initialize;
import xtr.mailing.renderizar.Render;


public class Funcoes {

	Session session;
	JSONObject dbList;
	JSONObject viewList;

	public Funcoes ( Session s ) {
		this.session = s;
	}

	public void setDbList( JSONObject dblista ) {
		this.dbList = dblista;
	}

	public void setViewList( JSONObject viewlista ) {
		this.viewList = viewlista;
	}


	/*
	 * %REM Function hasJob Description: Comments for Function %END REM
	 */
	public boolean hasJob ( Database dbJb, View vjb, String xtrcod ) {
		try {
			Document doc = vjb.getDocumentByKey(xtrcod, true);
			if (doc != null) return true;
			else return false;
		} catch (NotesException e) {
			e.printStackTrace();
			return false;
		}
	}

	/**
	 * <strong>incrementQtd</strong> - Incrementa a quantidade de um campo no documento.
	 * @param doc - Documento que será incrementado o campo.
	 * @param campo - Nome do campo.
	 */
	public void incrementQtd ( Document doc, String campo ) {
		try {
			double qtd;
			qtd = doc.getItemValueDouble(campo);
			qtd = qtd + 1;
			doc.replaceItemValue(campo, qtd);
			doc.save(true, false);
		} catch ( NotesException e ) {
			e.printStackTrace();
		}
	}

	/*
	 * %REM Function lockDocument Description: Função para travar documento
	 * evitando que outros peguem para fazer alguma tarefa %END REM
	 */

	public void lockDocument( Document job ) {
		int lock;
		try {
			if (job != null) {
				if (!job.hasItem("lock")) {
					job.replaceItemValue("lock", Integer.valueOf(1));
				} else if (job.getFirstItem("lock").getType() == 768) {
					lock = Integer.parseInt(job.getItemValue("lock").get(0)
							.toString());
					if (lock == 0) {
						job.replaceItemValue("lock", Integer.valueOf(1));
					}
				} else {
					job.replaceItemValue("lock", Integer.valueOf(1));
				}
			}
		} catch (NotesException e) {
			e.printStackTrace();
		}
	}

	/*
	 * %REM Function unlockDocument Description: Função para destravar documento
	 * %END REM
	 */

	public void unlockDocument( Document job ) {
		try {
			if (job != null) {
				job.replaceItemValue("lock", new Integer(0));
			}
		} catch (NotesException e) {
			e.printStackTrace();
		}
	}

	/*
	 * %REM Function isLockedDocument Description: Função que verifica se o
	 * documento está bloqueado para acesso %END REM
	 */

	public boolean isLockedDocument( Document job ) {
		boolean locked = false;
		int lock;
		try {
			if (job != null) {
				lock = Integer.parseInt(job.getItemValue("lock").get(0)
						.toString());
				if (!job.hasItem("lock")) {
					if (job.getFirstItem("lock").getType() == 768) {
						if (lock == 1)
							locked = true;
					}
				}
			}
			return locked;
		} catch (NotesException e) {
			e.printStackTrace();
			return false;
		}
	}

	/**
	 * <strong>createDocumentCampanha</strong> - Cria o documento inicial da
	 * campanha de envio de email.
	 * @param job - Document com as informações para criar o documento de envio.
	 * @param listdb - Lista de databases.
	 */
	public void createDocumentCampanha() {

		/*
		 * Function createDocumentCampanha( listaDb ) As NotesDocument Dim doc
		 * As NotesDocument Dim modelo, docSelecionados Dim docModelo As
		 * NotesDocument Dim vModelo As NotesView Dim ref_modeloemail
		 * 
		 * 'setando o nome da function nameFunction = "CreateDocumentCampanha"
		 * 
		 * On Error GoTo erro
		 * 
		 * 'modelo = job.codModelo(0) ref_modeloemail = job.ref_modeloemail(0)
		 * 
		 * 'docSelecionados = jsonRead.Parse( job.docSelecionados(0) ).items Set
		 * doc = listaDB.getItem( "emailcampanha" ).createDocument() Set vModelo
		 * = listaDB.getItem( "modeloemail" ).getView( "MODELOEMAIL-cod" ) Set
		 * docModelo = vModelo.Getdocumentbykey( modelo, true )
		 * 
		 * With doc .xtr_cod = entidadeNomeVersao( "emailcampanha" ) + "-" +
		 * doc.Universalid .ref_modeloemail = ref_modeloemail '.ref_modeloemail
		 * = modelo '.ref_emailcampanha = job.ref_emailcampanha(0) .nomeCampanha
		 * = job.nomecampanha(0) .modelo = modelo .nomemodelo =
		 * docModelo.descricaomailform(0) .qtdContatos = 0 .qtdEmails = 0
		 * .qtdBlacklist = 0 .qtdLog = 0 .qtdEnvios = 0 .qtdClicks = 0
		 * .dataCriacao = Now End With Call doc.save( true, true ) Set
		 * createDocumentCampanha = doc Exit Function erro: nameFunction =
		 * nameFunction + ": error na linha" + Str(Err) + ": " + Error$ End
		 * Function
		 */
	}

	/**
	 * <strong>createDocumentFila</strong> - Cria o documento na fila de email com todos os 
	 * parâmetros necessários para o envio 
	 * @param job - Documento do disparo com as informações para criar a fila de email.
	 * @param docPrincipal - Documento de contato.
	 * @param docModelo - Documento do modelo de email.
	 * @param docCampanha - Documento da campanha de envio.
	 * @param docMetrica - Documento das metricas de envio de email.
	 */
	public void createDocumentFila(Document job, Document docPrincipal, Document docModelo,
			Document docCampanha, Document docMetrica) {

		Document docFila;
		Document person;
		Document docAuxModelo;
		JSONObject docsRelacionamentos; 
		JSONObject listasCampos;

		try {
			//se documento for vazio sai da função 
			if ( docPrincipal == null ) return;

			JSONObject entidadesArquivos = 	new JSONObject(job.getItemValueString("entidadesArquivos"));
			JSONArray contasSmtp = 			new JSONArray(job.getItemValueString("contassmtp")); 
			JSONArray entidadesBasicas = 	new JSONArray(job.getItemValueString("entidadesBasicas"));
			JSONArray entidadesApp = 		new JSONArray(job.getItemValueString("entidadesApp"));

			//se tiver filtro e não atender a todos os filtros sai da função
			//comentado para erro na hora que remove item '
			//if Not  filtroSegmentacao(docPrincipal) ) } Exit Function

			//adicionando qtd de contatos na campanha e nas metricas Call
			incrementQtd(docCampanha, "qtdcontatos"); 
			incrementQtd(docMetrica, "qtdcontatos");

			//pegando copia do modelo de email para usa como auxiliar Set
			docAuxModelo = session.getCurrentDatabase().createDocument(); 
			docModelo.copyAllItems(docAuxModelo, true);

			//lista de documentos de relacionamentos 
			docsRelacionamentos = xtrListaRelacionamentos(job, docPrincipal);
			//lista de campos
			listasCampos = xtrListaCampos(job, docPrincipal, docsRelacionamentos); 

			//criando o documento person que vai para renderização das variaveis
			person = session.getCurrentDatabase().createDocument();
			//person.replaceItemValue("nomecampanha", nomecampanha ); 
			person.replaceItemValue( "ref_modeloemail", docModelo.getItemValueString("xtr_cod") );
			person.replaceItemValue( "ref_pessoa", docPrincipal.getItemValueString("xtr_cod") );
			person.replaceItemValue( "ref_metricaemail",docMetrica.getItemValueString("xtr_cod") );
			person.replaceItemValue( "ref_emailcampanha", job.getItemValue("ref_emailcampanha(0)") );
			person.replaceItemValue( "ref_blacklistcadastro", job.getItemValue("ref_blacklistcadastro") );

			//if job.blacklist(0) <> "" And job.blacklistxtrcod(0) <> "" } '
			//person.blacklistxtrcod = job.blacklistxtrcod(0) '}

			//passando os campos para o person 
			if ( listasCampos.length() > 0 ) {
				Iterator<String> it = listasCampos.keys();
				while ( it.hasNext() ) { 
					String lc = it.next();
					Item ni = (Item) listasCampos.get(lc);
					person.copyItem(ni, lc); 
				}
			}

			//criando o campo primnome 
			if ( person.hasItem("nome") ) {
				person.replaceItemValue( "primnome", person.getItemValueString("nome") ); 
			}

			//se não tiver email sai da função 
			if ( !hasEmail(person) ) return;

			incrementaQtdEmails(person , docCampanha);
			incrementaQtdEmails(person , docMetrica);

			filtroBlacklist(job, person, docCampanha, docMetrica); 

			//verifica se ainda tiver email depois do filtro de blacklist 
			if ( !hasEmail(person) ) return;

			filtroLogEmail(job, person, docCampanha, docMetrica);

			//verifica se existe email depois do filtro de log 
			if ( !hasEmail(person) ) return;

			//passando os arquivos de banco de dados usandos na pre renderização
			//dos modelo de email 
			Iterator<String> it = entidadesArquivos.keys();
			while ( it.hasNext() ) {
				String ite = it.next();
				JSONArray as = entidadesArquivos.getJSONArray(ite);	
				person.replaceItemValue(ite + "_Servidor", as.getString(0) ); 
				person.replaceItemValue( ite +	"_Arquivo", as.getString(1) ); 
			}

			//renderizando o modelo de email para cada documento selecionado
			//Render render = new Render(session, dbList);
			//render.setExecutDoc(person);
			//xtrRenderiza("ModeloMailForm", docAuxModelo, person, listasDB);
			//xtrRenderiza("AssuntoMailForm", docAuxModelo, person, listasDB);
			

			//criando documento da fila de email
			docFila = ((Database) dbList.get("filaemail")).createDocument(); 
			docFila.replaceItemValue( 
					"xtr_cod", 
					entidadeNomeVersao(entidadesApp, "filaemail") + "-" + docFila.getUniversalID()
			);

			docFila.replaceItemValue("status" ,"AGUARDANDO");
			//data de criação e usuario
			docFila.replaceItemValue( "xtr_criado_data", docFila.getCreated() ); 
			docFila.replaceItemValue( "xtr_criado_usuario", job.getItemValueString("xtr_criado_usuario") );

			//parametros do cabeçalho de envio e corpo do email 
			docFila.replaceItemValue( "rcptto", person.getItemValue("email") );

			String emailEnvio; 
			String nomeEnvio; 
			emailEnvio = docAuxModelo.getItemValueString("addrespostamailform"); 
			nomeEnvio = docAuxModelo.getItemValueString("nomerespostamailform"); 

			if ( emailEnvio.equals("") ) {
				docFila.replaceItemValue("mailfrom", emailEnvio); 
				docFila.replaceItemValue("from", nomeEnvio + " <" + emailEnvio + ">"); 
				docFila.replaceItemValue("namefrom", nomeEnvio.trim()); 
			} else {
				docFila.replaceItemValue("mailfrom", job.getItemValueString("xtr_criado_usuario"));
				docFila.replaceItemValue("from", job.getItemValueString("xtr_criado_usuario"));
				docFila.replaceItemValue("namefrom",""); 
			}

			docFila.replaceItemValue("to", person.getItemValueString("email"));

			//docFila.copyto = "jonatanraimir@consiste.com.br"
			//docFila.blindcopyto = "jonatanraimir@consiste.com.br"

			docFila.replaceItemValue("subject", docAuxModelo.getItemValueString("assuntoMailForm")); 
			docFila.replaceItemValue("body", returnTextOfMimeOrRichText(docAuxModelo, "modeloMailForm"));

			//parâmetros da conexão com o servidor 
			Vector<String> user = new Vector<String>(); 
			Vector<String> senha = new Vector<String>(); 
			Vector<String> server = new Vector<String>(); 
			Vector<Integer> port = new Vector<Integer>(); 
			for( Object cont : contasSmtp ) {
				JSONObject conta = new JSONObject(cont.toString());
				user.add(conta.getString("usuario"));
				senha.add(conta.getString("senha")); 
				server.add(conta.getString("servidor")); 
				port.add( Integer.valueOf(conta.getInt("porta")) );
			}

			docFila.replaceItemValue("usersmtp", user); 
			docFila.replaceItemValue("passwordsmtp", senha); 
			docFila.replaceItemValue("server", server); 
			docFila.replaceItemValue("port", port); 
			docFila.replaceItemValue("autenticado", job.getItemValueDouble("autenticado"));
			docFila.replaceItemValue("status", "AGUARDANDO");

			//parâmetros do modelo 
			docFila.replaceItemValue("logservidor", ((Database) dbList.get("logemail")).getServer());
			docFila.replaceItemValue("logarquivo", ((Database) dbList.get("logemail")).getFilePath()); 
			docFila.replaceItemValue("nomecampanha", job.getItemValueDouble("nomecampanha"));

			docFila.replaceItemValue("metricaservidor", ((Database) dbList.get("metricaemail")).getServer());
			docFila.replaceItemValue("metricaarquivo", ((Database) dbList.get("metricaemail")).getFilePath());
			docFila.replaceItemValue("metricacod", docMetrica.getItemValueString("xtr_cod"));

			docFila.replaceItemValue("emailcampanhaservidor", ((Database) dbList.get("emailcampanha")).getServer());
			docFila.replaceItemValue("emailcampanhaarquivo", ((Database) dbList.get("emailcampanha")).getFilePath());

			docFila.replaceItemValue("logmodeloemail", job.getItemValueString("ref_modeloemail"));
			docFila.replaceItemValue("ref_contato", docPrincipal.getItemValueString("xtr_cod")); 
			docFila.replaceItemValue("ref_modeloemail", job.getItemValueString("ref_modeloemail"));
			docFila.replaceItemValue("ref_Segmentacao", job.getItemValueString("ref_Segmentacao"));
			docFila.replaceItemValue("ref_blacklistcadastro", job.getItemValueString("ref_blacklistcadastro"));
			docFila.replaceItemValue("ref_emailcampanha", job.getItemValueString("ref_emailcampanha"));
			docFila.replaceItemValue("ref_metricaemail", job.getItemValueString("ref_metricaemail"));

			docFila.replaceItemValue("nomecontato", person.getItemValueString("nome") + " " +
					person.getItemValueString("sobrenome")
			);

			docFila.save(true, false);
			docCampanha.save(true, false);
			docMetrica.save(true, false);
		} catch ( NotesException e ) {
			e.printStackTrace();
		}
	}

	/**
	 * <strong>createDocMetrica</strong> - Cria o documento inicial das
	 * metricas de envio de email.
	 * @param job - Document com as informações para criar o documento de envio.
	 * @param listdb - Lista de databases.
	 */
	public Document createDocumentMetrica(Document job, JSONObject listdb) {
		try {
			View vModelo;
			Document doc;
			Document docModelo;
			JSONArray entidadesApp;
			entidadesApp = new JSONArray(job.getItemValueString("entidadesApp"));
			String entidade = entidadeNomeVersao(entidadesApp, "metricaemail");

			vModelo = ((Database) listdb.get( "modeloemail" )).getView("MODELOEMAIL-cod");
			doc = ((Database) listdb.get("metricaemail") ).createDocument();
			doc.replaceItemValue("xtr_cod",  entidade + "-" + doc.getUniversalID());
			docModelo = vModelo.getDocumentByKey( job.getItemValueString("ref_modeloemail"), true );
			doc.replaceItemValue( "ref_modeloemail", docModelo.getItemValueString("xtr_cod") );
			doc.replaceItemValue( "ref_emailcampanha", job.getItemValueString("ref_emailcampanha") );
			doc.replaceItemValue("nomeCampanha", job.getItemValueString("nomecampanha"));
			//doc.replaceItemValue("modelo", job.getItemValueString("nomemodelo"));
			//doc.replaceItemValue("nomemodelo = docModelo.descricaomailform(0);
			Integer num = new Integer(0);
			doc.replaceItemValue("qtdContatos", num);
			doc.replaceItemValue("qtdEmails", num);
			doc.replaceItemValue("qtdBlacklist", num);
			doc.replaceItemValue("qtdLog", num);
			doc.replaceItemValue("qtdEnvios", num);
			doc.replaceItemValue("qtdClicks", num);
			doc.replaceItemValue("dataCriacao", doc.getCreated());		
			doc.save(true, true);
			return doc;
		} catch ( NotesException e ) {
			e.printStackTrace();
			return null;
		}
	}

	/**
	 * <strong>entidadeNomeVersao</strong>
	 * @param entidade - String com o nome da entidade que está buscando.
	 * @param entidades - JSONArray com a lista de entidades para o envio de email.
	 * @return Retorna o nome e a versao da entidade passada como parametro.
	 */
	public String entidadeNomeVersao( JSONArray entidades, String entidade ) {
		Iterator<Object> e = entidades.iterator();
		String ent = "";
		while ( e.hasNext() ) {
			ent = (String) e.next();
			if ( ent.indexOf(entidade) != -1 ) {
				return ent;
			}
		}	
		return "";
	}

	/*
	 * %REM Function filtroSegmentacao Description: Essa função faz uma busca
	 * dos segmento dentro do documento passado como parâmetro %END REM
	 */
	public void filtroSegmentacao() {
		/*
		 * Function filtroSegmentacao( docPessoa As NotesDocument ) As Boolean
		 * Dim dbSeg As NotesDatabase Dim vSeg As NotesView Dim docSegmentacao
		 * As NotesDocument, docSeg As NotesDocument Dim flag As Boolean
		 * 
		 * Dim docRel As NotesDocument Dim copyDocPessoa As NotesDocument Dim
		 * camposRel(7), camposSeg(12), notPessoa(3) Dim items, itemsConvert Dim
		 * Segmentacaos, valsPerson
		 * 
		 * 'setando o nome da function nameFunction = "filtroSegmentacao"
		 * 
		 * On Error GoTo erro
		 * 
		 * Set vSeg = listasView.getItem("segmentacao") Set docSegmentacao =
		 * vSeg.Getdocumentbykey( ref_Segmentacao, true )
		 * 
		 * 'retorna true se o usuário não passou um segmento no formulário if
		 * docSegmentacao != null } filtroSegmentacao = true Exit Function }
		 * 
		 * 'campos necessários em uma pessoa camposRel(0) = "sexo" camposRel(1)
		 * = "categorias" camposRel(2) = "religiao" camposRel(3) = "tabagismo"
		 * camposRel(4) = "cidade" camposRel(5) = "tiposangue" camposRel(6) =
		 * "cargo" camposRel(7) = "nivelestrategico"
		 * 
		 * 'campos necessários em uma segmento camposSeg(0) = "sexo"
		 * camposSeg(1) = "tabagismo" camposSeg(2) = "tiposangue" camposSeg(3) =
		 * "excluir_categorias" camposSeg(4) = "incluir_categorias" camposSeg(5)
		 * = "excluir_religiao" camposSeg(6) = "incluir_religiao" camposSeg(7) =
		 * "excluir_cidade" camposSeg(8) = "incluir_cidade" camposSeg(9) =
		 * "excluir_cargo" camposSeg(10) = "incluir_cargo" camposSeg(11) =
		 * "excluir_nivelestrategico" camposSeg(12) = "incluir_nivelestrategico"
		 * 
		 * 'campos necessários para apagar o segmento do filtro se o tipo de
		 * documento não for um vinculo notPessoa(0) = "excluir_cargo"
		 * notPessoa(1) = "incluir_cargo" notPessoa(2) =
		 * "excluir_nivelestrategico" notPessoa(3) = "incluir_nivelestrategico"
		 * 
		 * 
		 * flag = false Set copyDocPessoa =
		 * session.currentDatabase.createDocument Set docSeg =
		 * session.Currentdatabase.CreateDocument Call
		 * docPessoa.Copyallitems(copyDocPessoa, true) Call
		 * docSegmentacao.Copyallitems(docSeg, true)
		 * 
		 * 'copia os campos de relacionamento se houver if
		 * docPessoa.hasItem("ref_pessoa") And
		 * docPessoa.getItemValue("ref_pessoa")(0) <> "" } Set docRel =
		 * listasView
		 * .getItem("pessoa").getDocumentByKey(docPessoa.getItemValue(
		 * "ref_pessoa" )(0)) if Not docRel != null } Call
		 * docRel.Copyallitems(copyDocPessoa, true) } }
		 * 
		 * 'retira os campos do documento de pessoa que não faz parte do filtro
		 * de segmento items = copyDocPessoa.Items ForAll item In items if
		 * IsNull(ArrayGetIndex(camposRel, item.name, 5)) } Call item.Remove()
		 * else itemsConvert = item.Value ForAll itc In itemsConvert itc =
		 * LCase(removeAcentos(itc)) 'retirando acentuação e colocando em
		 * minúsculo End ForAll item.Values = itemsConvert } End ForAll
		 * 
		 * 'retira os campos do documento de segmento que não faz parte dos
		 * filtros items = docSeg.Items ForAll item In items if
		 * IsNull(ArrayGetIndex(camposSeg, item.name, 5)) } Call item.Remove()
		 * else itemsConvert = item.Values ForAll itc In itemsConvert itc =
		 * LCase(removeAcentos(itc)) 'retirando acentuação e colocando em
		 * minúsculo End ForAll item.Values = itemsConvert } End ForAll
		 * 
		 * 'apaga alguns campos se o documento não for um vinculo if Not
		 * docPessoa.hasItem("ref_pessoa") or
		 * docPessoa.getItemValue("ref_pessoa")(0) = "" } items = docSeg.Items
		 * ForAll item In items if Not IsNull(ArrayGetIndex(notPessoa,
		 * item.name, 5)) } Call item.Remove() } End ForAll }
		 * 
		 * 
		 * items = docSeg.Items ForAll item In items if item.name = "sexo" Or
		 * item.name = "tabagismo" Or item.name = "tiposangue" } flag = false
		 * Segmentacaos = docSeg.getItemValue(item.name) valsPerson =
		 * copyDocPessoa.getItemValue(item.name)
		 * 
		 * ForAll segmento In Segmentacaos if segmento <> "" And segmento <>
		 * "todos" } if segmento = "nao informado" } if Not IsNull(
		 * ArrayGetIndex(valsPerson, "", 5) ) } if (UBound(valsPerson) + 1) < 2
		 * } flag = true } } else if Not IsNull( ArrayGetIndex(valsPerson,
		 * segmento, 5) ) } flag = true } else flag = true } End ForAll
		 * 
		 * if Not flag } filtroSegmentacao = false Exit Function } } End ForAll
		 * 
		 * 
		 * 'Excluir contatos com estes segmentos 'if Not
		 * filtroSegmentacaoExcluir( copyDocPessoa, docSeg ) } '
		 * filtroSegmentacao = false ' Exit Function '}
		 * 
		 * 'Incluir contatos que tenha esses segmentos 'if Not
		 * filtroSegmentacaoIncluir( copyDocPessoa, docSeg ) } '
		 * filtroSegmentacao = false ' Exit Function '}
		 * 
		 * filtroSegmentacao = flag Exit Function erro: nameFunction =
		 * nameFunction + ": error na linha" + Str(Err) End Function
		 */
	}


	/**
	 * <strong>filtraEmail</strong> - Essa função faz um filtro pelo
	 * tipo de check para saber se já foi enviado essa campanha para o email.
	 * @param job - Documento de job com todas a informações
	 * @param docEmails - Documento da pessoa com os emails a ser filtrado.
	 * @param docCampanha - Documento de contagem da campanha.
	 * @para docMetrica - Documento de contagem das métricas.
	 */
	@SuppressWarnings("unchecked")
	public void filtroLogEmail(Document job, Document docEmails, Document docCampanha, 
			Document docMetrica) {
		try {
			Database dbLog = null;
			View vLog = null; 
			Document docLog; 
			String check, email;
			String ref_modelo, ref_campanha;
			Vector emails, newEmails, src;
			double qtdMet, qtdCamp;

			src = new Vector();
			newEmails = new Vector();
			dbLog = (Database) dbList.get("logemail");
			check = job.getItemValueString("checkEnvio");
			check = check.toLowerCase();
			//se o banco de log não existir sai da função 
			if (dbLog == null) return;

			//verificando se o tipo de check é diferente de sem check
			if ( check.equals("nenhum") ) return;

			ref_modelo = job.getItemValueString("ref_modeloemail"); 
			ref_campanha = docCampanha.getItemValueString("xtr_cod");

			qtdCamp = docCampanha.getItemValueDouble("qtdLog"); 
			qtdMet = docMetrica.getItemValueDouble("qtdLog"); 
			emails = docEmails.getItemValue("email");

			if ( check.equals("campanha") ) {
				vLog = dbLog.getView("LOGEMAIL-emailcampanha-email");
				src.add(ref_campanha);
			} 
			else if ( check.equals("modelo") ) {
				vLog = dbLog.getView("LOGEMAIL-modeloemail-email"); 
				src.add(ref_modelo); 
			}

			for( Object e : emails ) {
				email = e.toString().trim(); 
				if ( !email.equals("") ) {
					src.add(email);
					docLog = vLog.getDocumentByKey(src, true);
					//se não encontrou emails então não houve filtro do email 
					if ( docLog == null ) {
						newEmails.add(email); 
					} else { 
						qtdCamp += 1; 
						qtdMet += 1; 
					}
				}	
			}

			//Atualizando os valores nas metricas e salvando 
			docCampanha.replaceItemValue("qtdLog", qtdCamp); 
			docMetrica.replaceItemValue("qtdLog", qtdMet);

			docCampanha.save(true, false); 
			docMetrica.save(true, false);
			docEmails.replaceItemValue("email", newEmails);
		} catch ( NotesException e) {
			e.printStackTrace();
		}
	}

	/**
	 * <strong>hasEmail</strong> -Essa função verifica se
	 * existe os campos email e/ou outrosEmails, se é diferente de
	 * vazio então junta em um mesmo campo chamado email e retorna um boolean
	 * para saber se existe email no documento selecionado.
	 * @param doc - Documento da contato com o email.
	 */
	@SuppressWarnings("unchecked")
	public boolean hasEmail(Document doc) {
		try {
			String em;
			boolean flag = false;
			Vector<Object> email = doc.getItemValue("email");
			email.addAll( doc.getItemValue("outrosEmail") );
			email.addAll( doc.getItemValue("outrosEmails") );

			Vector emails = new Vector();


			for( int x = 0; x < email.size(); x++ ) {
				em = email.get(x).toString();
				if ( !em.equals("") ) {
					if ( !emails.contains(em) ) emails.add(em); 
				}
			}


			if ( !emails.isEmpty() && !emails.get(0).equals("") ) {  
				flag = true;
			}

			doc.removeItem("email"); 
			doc.removeItem("OutrosEmails");
			doc.removeItem("OutrosEmail");
			doc.replaceItemValue("email", emails); 
			return flag;
		} catch ( NotesException e ) {
			return false;
		}	
	}

	/**
	 * <strong>incrementaQtdEmails</strong> - Incrementa a quantidade de
	 * emails no documento de métrica ou de campanha.
	 * @param docPessoa - Documento da pessoa com os emails.
	 * @param doc - Documento da métrica ou campanha.
	 */
	@SuppressWarnings("unchecked")
	public void incrementaQtdEmails( Document docPessoa, Document doc) {
		try {
			double count, qtdEmails;
			Vector emails = docPessoa.getItemValue("email"); 
			qtdEmails = doc.getItemValueDouble("qtdEmails"); 
			count = emails.size(); 
			qtdEmails += count;
			doc.replaceItemValue("qtdEmails", qtdEmails);
			doc.save(true, false);
		} catch ( NotesException e) {
			e.printStackTrace();
		}
	}

	/*
	 * %REM Function filtroSegmentacaoExcluir Description: excluir as pessoas
	 * que tem os seguintes segmentos como parâmetro %END REM
	 */
	public void filtroSegmentacaoExcluir() {
		/*
		 * Function filtroSegmentacaoExcluir( docPessoa As NotesDocument, docSeg
		 * As NotesDocument ) As Boolean 'todos os items Dim items Dim
		 * Segmentacaos, valsPerson, flag
		 * 
		 * 'setando o nome da function nameFunction = "filtroSegmentacaoExcluir"
		 * 
		 * On Error GoTo erro
		 * 
		 * flag = false items = docSeg.Items
		 * 
		 * ForAll item In items if InStr(1, item.name, "excluir_", 5) = 1 } flag
		 * = false Segmentacaos = item.values valsPerson =
		 * docPessoa.getItemValue(StrRight(item.name, "excluir_"))
		 * 
		 * ForAll segmento In Segmentacaos if segmento <> "" } if Not IsNull(
		 * ArrayGetIndex(valsPerson, segmento, 5) ) } filtroSegmentacaoExcluir =
		 * false Exit Function else flag = true } else flag = true } End ForAll
		 * 
		 * if Not flag } filtroSegmentacaoExcluir = false Exit Function } } End
		 * ForAll
		 * 
		 * filtroSegmentacaoExcluir = flag Exit Function erro: nameFunction =
		 * nameFunction + ": error na linha" + Str(Err) End Function
		 */
	}

	/**
	 * <strong>xtrListaCampos</strong> - Esse função pega os campos do
	 * documento principal retornando uma lista com com objetos notesitem.
	 * @param job - Documento job para busca de informações do envio.
	 * @param doc - Documento do contato que será enviado.
	 * @param docRel - Lista de documentos de relacionamento do contato.
	 * @return JSONObject - Lista de NotesItem para inserir no documento de informações de envio.
	 */
	public JSONObject xtrListaCampos( Document job, Document doc, JSONObject docRel) {
		JSONObject listaCampos = new JSONObject();
		try { 
			String relCampo[], rel, lc, campo;
			Item campoValue;

			JSONObject campos;
			campos = new JSONObject(job.getItemValueString("campos"));
			Iterator<String> ic = campos.keys();

			//saindo da função e retornando se o tamanho for 0.
			if ( campos.length() == 0 ) {
				return new JSONObject();
			}

			while ( ic.hasNext() ) {
				lc = ic.next();
				if ( lc.indexOf(".") != -1 ) { 
					relCampo = lc.split("."); 
					rel = "ref_" + relCampo[0].toLowerCase(); 
					campo = relCampo[1]; 
					campoValue = ((Document) docRel.get(rel)).getFirstItem(campo); 
				} else {
					campoValue = doc.getFirstItem(lc);
				}

				//salva se o item existe no documento 
				if ( campoValue != null ) { 
					listaCampos.put(lc, campoValue ); 
				} 
			}
			return listaCampos;
		} catch ( NotesException e ) {
			e.printStackTrace();
			return listaCampos;
		}
	}

	/**
	 * <strong>xtrListaDatabase</strong><br>
	 * 
	 * @param lista - JSONObject com nome da base: [arquivo, servidor].
	 * @return Retorna uma lista de objetos NotesDatabase não nulos.
	 */
	public JSONObject xtrListaDatabase(JSONObject lista) {
		Database db;
		JSONArray param;
		JSONObject listdb = new JSONObject();
		String bd;

		try {
			Iterator<String> keys = lista.keys();

			while (keys.hasNext()) {
				bd = keys.next();
				param = lista.getJSONArray(bd);
				db = session.getDatabase(param.get(0).toString(), param.get(1)
						.toString(), false);
				if (db != null) {
					if (db.isOpen())
						listdb.put(bd, db);
				}
			}
			dbList = listdb;
			return listdb;
		} catch (NotesException e) {
			e.printStackTrace();
			return listdb;
		}
	}

	/**
	 * <strong>xtrListView</strong>
	 * 
	 * @param listdb
	 *            - Lista de NotesDatabase.
	 * @return Lista de view relacionadas a lista de NotesDatabase.
	 */
	public JSONObject xtrListaView(JSONObject listdb) {
		String bd;
		Database db;
		View view;
		JSONObject listView = new JSONObject();

		try {
			Iterator<String> keys = listdb.keys();

			while (keys.hasNext()) {
				bd = keys.next();
				db = (Database) listdb.get(bd);
				if (db != null) {
					if (db.isOpen()) {
						view = db.getView(bd.toUpperCase() + "-cod");
						if (view != null) listView.put(bd, view);
					}
				}
			}
			viewList = listView;
			return listView;
		} catch (NotesException e) {
			e.printStackTrace();
			return listView;
		}
	}

	/**
	 * <strong>xtrListaRelacionamentos</strong> - Essa função cria os
	 * documentos de relacionamentos se existirem relacionamentos.
	 * @param job - Documento job para busca de informações do envio.
	 * @param docP - Documento do contato que será enviado.
	 * @return JSONObject - Lista de documentos de relacionamento do contato.
	 */
	public JSONObject xtrListaRelacionamentos( Document job, Document docP ) {
		JSONObject lista = new JSONObject();

		try {
			Database db; 
			View v;
			Document doc; 
			String cod, ent[], entidade;
			JSONArray relacionamentos;

			if ( !job.getItemValueString("relacionamentos").equals("{}") && 
					!job.getItemValueString("relacionamentos").equals("") ) { 

				relacionamentos = new JSONArray(job.getItemValueString("relacionamentos"));

				if ( relacionamentos.isNull(0) || !(relacionamentos.length() > 0) ) { 
					return new JSONObject(); 
				}

				for ( Object r : relacionamentos ) {
					if ( !(r.toString().equals("")) ) {
						cod = docP.getItemValueString(r.toString()); 
						ent = cod.split("-"); 
						entidade = ent[0]; 
						db = (Database) dbList.get(entidade);
						v = db.getView(entidade.toUpperCase() + "-cod"); 
						doc = v.getDocumentByKey(cod, true);
						if (doc != null) { 
							lista.put(r.toString(), doc);
						}
					}	
				} 
			}
			return lista;
		} catch( NotesException e ) {
			e.printStackTrace();
			return lista;
		}
	}




	/*
	 * %REM Function filtroSegmentacaoIncluir Description: inclui as pessoas que
	 * tem os seguintes segmentos como parâmetro %END REM
	 */
	public void filtroSegmentacaoIncluir() {
		/*
		 * Function filtroSegmentacaoIncluir( docPessoa As NotesDocument, docSeg
		 * As NotesDocument ) As Boolean
		 * 
		 * 'todos os items Dim items Dim Segmentacaos, valsPerson, flag
		 * 
		 * flag = false items = docSeg.Items
		 * 
		 * ForAll item In items if InStr(1, item.name, "incluir_", 5) } flag =
		 * false Segmentacaos = item.values valsPerson =
		 * docPessoa.getItemValue(StrRight(item.name, "incluir_"))
		 * 
		 * ForAll segmento In Segmentacaos if segmento <> "" } if IsNull(
		 * ArrayGetIndex(valsPerson, segmento, 5) ) } filtroSegmentacaoIncluir =
		 * false Exit Function else flag = true } else flag = true } End ForAll
		 * 
		 * if Not flag } filtroSegmentacaoIncluir = false Exit Function } } End
		 * ForAll
		 * 
		 */
	}

	/**
	 * <strong>filtroBlacklist</strong> - Faz o filtro dos email na
	 * black list e retira os da lista de email a ser enviada.
	 * @param job - Documento de job com todas a informações
	 * @param docEmails - Documento da pessoa com os emails a ser filtrado.
	 * @param docCampanha - Documento de contagem da campanha.
	 * @para docMetrica - Documento de contagem das métricas.
	 */
	@SuppressWarnings("unchecked")
	public void filtroBlacklist(Document job, Document docEmails,
			Document docCampanha, Document docMetrica) {

		try {
			Database dbBl;
			View vBl;
			Document doc; 
			String ref_blacklist;
			String email;
			Vector src;
			Vector emails, newEmails;
			Double qtdCamp, qtdMet;

			src = new Vector();
			newEmails = new Vector();
			ref_blacklist = job.getItemValueString("ref_blacklistcadastro");

			//se o usuário não informou a blacklist não faz o filtro 
			if ( ref_blacklist.equals("") ) return;

			qtdCamp = docCampanha.getItemValueDouble("qtdBlacklist"); 
			qtdMet = docMetrica.getItemValueDouble("qtdBlacklist");

			emails = docEmails.getItemValue("email"); 
			dbBl = (Database) dbList.get("blacklist");
			vBl = dbBl.getView("BLACKLIST-blacklist-email");

			src.add(ref_blacklist);

			for( Object e : emails ) {
				email = ((String) e).trim(); 
				if ( !email.equals("") ) { 
					src.add(email);
					doc = vBl.getDocumentByKey(src, true); 
					if ( doc == null ) {
						newEmails.add(email); 
					}
					else { 
						qtdCamp = qtdCamp + 1;
						qtdMet = qtdMet + 1;
					}
				}
			}

			//Atualizando os valores nas metricas e salvando
			docCampanha.replaceItemValue("qtdBlacklist", qtdCamp);
			docMetrica.replaceItemValue("qtdBlacklist", qtdMet);
			docCampanha.save(true, false);
			docMetrica.save(true, false);
			docEmails.replaceItemValue("email", newEmails);
		} catch ( NotesException e ) {
			e.printStackTrace();
		}
	}

	/**
	 * <strong>returnTextOfMimeOrRichText</strong> - Retorna o texto do mime ou do campo 
	 * texto do documento do modelo de email
	 * @param doc - Documento de onde vai pegar o campo MIME.
	 * @param campo - Nome do campo que vai pegar o texto
	 */
	public String returnTextOfMimeOrRichText( Document doc, String campo) {
		try {
			int type; 
			String texto = "";
			Item item;

			session.setConvertMIME(false);

			item = doc.getFirstItem(campo);
			type = item.getType();

			if ( type == Item.TEXT ) {
				texto = item.getValueString(); 
			}
			else if ( type == Item.RICHTEXT ) { 
				texto = ((RichTextItem) doc.getFirstItem(campo)).getUnformattedText();
			}
			else if ( type == Item.MIME_PART ) { 
				MIMEEntity mime = doc.getMIMEEntity(campo);
				while( mime != null ) {
					if ( (mime.getContentType() + "/" + mime.getContentSubType()).equals("text/html") ) {
						texto += mime.getContentAsText();
					}
					mime = mime.getNextEntity();
				}
			}
			return texto;
		} catch ( NotesException e ) {
			e.printStackTrace();
			return "";
		}
	}	
}
