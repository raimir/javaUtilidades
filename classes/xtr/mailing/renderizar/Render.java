package xtr.mailing.renderizar;

import java.util.ArrayList;
import java.util.Vector;

import lotus.domino.Agent;
import lotus.domino.Database;
import lotus.domino.Document;
import lotus.domino.Item;
import lotus.domino.MIMEEntity;
import lotus.domino.MIMEHeader;
import lotus.domino.NotesException;
import lotus.domino.NotesFactory;
import lotus.domino.RichTextItem;
import lotus.domino.Session;
import lotus.domino.Stream;
import lotus.domino.View;
import xtr.util.StringUtil;
import org.json.JSONArray;
import org.json.JSONObject;
import org.json.JSONException;

/**
Library xtrRederizacao
Created 07/03/2016 by Jonatan Raimir Santos da Silva/CONSISTE
Description: Essa biblioteca tem processos de renderização de variaveis de email 

Option Public
Use "htmlmaillib2"
 */

public class Render {
	private Session session;
	String executeResult;
	String executeResult2;
	Document executeDoc;
	Document executeDoc2;
	private Stream stream;
	private JSONObject dbList;
	
	public Render( Session session, JSONObject listasdb ) {
		this.dbList = listasdb;
		this.session = session;
	}
	
	public Render( Session session ) {
		this.session = session;
	}
	
	public Document Execute( String funcao, Document docPessoa ) {
		Document docresultado = null;
		try {
			docresultado = session.getCurrentDatabase().createDocument();
			docPessoa.copyAllItems(docresultado, true);
			docresultado.replaceItemValue("funcao3x3cut3", funcao);
			Agent ag = session.getCurrentDatabase().getAgent("abc");
			ag.runWithDocumentContext(docresultado);
			return docresultado;
		} catch( NotesException e ) {
			e.printStackTrace();
			try {
				docresultado = session.getCurrentDatabase().createDocument();
				docresultado.replaceItemValue("error3x3cut3", e.getMessage());
				return docresultado;
			} catch ( NotesException ex) {
				ex.printStackTrace();
				return docresultado;
			}
		}
	}

	/**
	public void removerCaracterEspeciais
	Description: Comments for public void
	 */
	public void removerCaracterEspeciais( String texto ) {
		if ( texto.indexOf("\\") != -1 ) {
			texto = texto.replace("\\", "\\\\");
		} else if ( texto.indexOf("\"") != -1 ) {
			texto = texto.replace("\"", "\\\"");
		}
	}

	/**
		public void replaceTextoModeloEmail
		Description: Modifica o valor do campo do modelo de email já configurado com o texto resolvido 
	 */
	public void replaceTextoModeloEmail( Document doc, String campo, String textoReplace ) {
		try {
			int type;
			MIMEEntity mime;
			MIMEHeader mimeHeader;
			
			session.setConvertMIME(false);
			stream = session.createStream();
			type = doc.getFirstItem(campo).getType();
			
			if ( type == Item.TEXT ) { 
				doc.replaceItemValue(campo, textoReplace);
			}	
			else if ( type == Item.MIME_PART ) { 
				doc.closeMIMEEntities(true , campo);
				doc.removeItem(campo);
				mime = doc.createMIMEEntity(campo);
				mimeHeader = mime.createHeader("Content-Type");
				mimeHeader.setHeaderVal("text/html; ");
				mimeHeader = mime.createHeader("MIME-Version");
				mimeHeader.setHeaderVal("1.0; ");
				mimeHeader.setParamVal("charset", "UTF-8;");

				stream.writeText(textoReplace, 5);
				mime.setContentFromText( stream, "text/html; charset=UTF-8;", 1725 );
				doc.closeMIMEEntities(true , campo);
				stream.close();
			}
		} catch (NotesException e) {
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
	
	public void setExecutDoc( Document doc ) {
		this.executeDoc2 = doc;
	}

	/**
	 * public void tratarFormula
	 * Description: Essa função pega o formula antes de renderizar o que tem
	 * dentro dela e substitui aspas dupla(") por chaves ({})
	 */
	public String tratarFormula(String texto) { 
		try { 
			String formula, copyTexto, abrirAspa, FecharAspa;

			//tratando as formulas
			copyTexto = "";
			while ( texto.indexOf("<%") != -1 ) {
				copyTexto += StringUtil.strLeft(texto, "<%");
				texto = StringUtil.strRight(texto, "<%");
				formula = StringUtil.strLeft(texto, "%>");

				while ( formula.indexOf("\"") != -1 ) {
					//primeira aspa
					abrirAspa = StringUtil.strLeft(formula, "\"");
					formula = formula.replace((abrirAspa + '"'),(abrirAspa + '{'));
					//segunda aspa
					if ( formula.indexOf("\"") != -1 ) {
						FecharAspa = StringUtil.strLeft( formula, "\"");
						formula = formula.replace( (FecharAspa + '"'),(FecharAspa + '}') );
					}
				}

				copyTexto += "<%" + formula + "%>";
				texto = StringUtil.strRight(texto, "%>");
			}
			copyTexto += texto;
			return copyTexto;
		} catch ( Exception e ) {
			e.printStackTrace();
			return texto;
		}	
	}

	/**
	 * public void xtrPreRenderiza
	 * Description: Essa função faz a pre renderização do modelo de email com variaveis
	 * que não necessita de informações do contato
	 */
	public void xtrPreRenderiza( String campo, Document docModelo ) {
		try {
			Document docVar;
			String variavel, qvar, funcaoLotus;
			String texto, prerend;
			String retorno = "";
			String erro = "";
			String subst = "";

			Database db = (Database) dbList.get("variavelemail");
			View v = db.getView("VARIAVELEMAIL-variavel");
			prerend = "";

			//pegando o texto mime
			texto = returnTextOfMimeOrRichText(docModelo, campo);
			//Renderizando as variáveis de COMPONENTES 
			texto = xtrRenderizaComponente(texto);
			//Renderizando as variáveis de LINKS 
			texto = xtrRenderizaLink(texto);
			
			//resolvendo as variavel de email
			while ( texto.indexOf("!*") != -1 ) {
				executeResult2 = "";
				variavel = "";

				prerend += StringUtil.strLeft(texto, "!*");
				texto = StringUtil.strRight(texto, "!*");
				qvar = (StringUtil.strLeft(texto, "*!")).toUpperCase();
				variavel = StringUtil.strLeft(texto, "*!");

				//texto = "!*" + texto;
				docVar = v.getDocumentByKey(qvar, false);

				if ( docVar != null ) {
					funcaoLotus = docVar.getItemValueString("lotusscript");
					Document docexe = session.getCurrentDatabase().createDocument();
					docexe = Execute(funcaoLotus, docexe);
					
					if ( docexe != null ) {
						retorno = docexe.getItemValueString("retorno3x3cut3");
						erro = docexe.getItemValueString("error3x3cut3");
					}
						
					if ( erro.equals("") && !retorno.equals("") ) {
						prerend += retorno;
					} else {
						prerend += "!*" + variavel + "*!";
					}	
				}
				else {
					prerend += "!*" + variavel + "*!";
				}
				subst = StringUtil.strRight(texto, "*!");
				texto = subst.equals("") ?  texto : subst;
			}
			prerend += texto;
			replaceTextoModeloEmail(docModelo, campo, prerend);
		} catch( NotesException e ) {
			e.printStackTrace();
		}
	}

	/**
	public void xtrRenderiza
	Description: Renderiza as variaveis que necessitam do documento pessoas
	 */
	public void xtrRenderiza( String campo, Document docModelo ) {
		try {
			Document docVar;
			String texto, text1, variavel, qvar;
			String formula, formulaCorr, funcaoLotus;

			Database db = (Database) dbList.get("variavelemail");
			View v = db.getView("VARIAVELEMAIL-variavel");
			texto = returnTextOfMimeOrRichText(docModelo, campo);

			//trocar aspa por chaves dentro da formula antes de resolver as variáveis e formulas
			texto = tratarFormula(texto);
			//Renderizando as variáveis de LINKS 
			texto = xtrRenderizaLink(texto);
			//Renderizando as variáveis de CLICKS 
			texto = xtrRenderizaClick(texto, executeDoc2);

			//resolvendo as variavel de email
			while ( texto.indexOf("!*") != -1 ) {
				executeResult = "";
				text1 = StringUtil.strRight(texto, "!*");
				variavel = StringUtil.strLeft(text1, "*!");
				qvar = (StringUtil.strLeft( text1, "*!" )).toUpperCase();

				docVar = v.getDocumentByKey(qvar, false);
				if ( docVar != null ) {
					funcaoLotus = docVar.getItemValueString("lotusscript");
					//Execute(funcaoLotus);
					Agent ag = session.getCurrentDatabase().getAgent("abc");
					//ag.runWithDocumentContext(arg0);
					texto = texto.replace("!*" + variavel + "*!", executeResult );
				} else {
					texto = texto.replace(("!*" + variavel + "*!"), "");
				}
			}


			//resolvendo as formulas
			while ( texto.indexOf("<%") != -1 ) {
				executeResult = "";
				text1 = StringUtil.strRight(texto, "<%");
				formula = StringUtil.strLeft(text1, "%>");
				//executeResult = Evaluate( Trim(formula), executeDoc )
				executeResult = xtrRenderizaFormula(formula); 
				texto = texto.replace("<%" + formula + "%>", executeResult);
			}

			replaceTextoModeloEmail(docModelo, campo, texto);
		} catch( NotesException e ) {
			e.printStackTrace();
		}
	}


	public String xtrRenderizaClick( String texto, Document docPessoa ) {
		try {
			Document doc;
			String codClick;
			Database db = (Database) dbList.get("clickcadastro");
			View view = db.getView("CLICKCADASTRO-variavel");
			
			//resolvendo as variavel de email
			while ( texto.indexOf("!*CLICK_") != -1 ) {
				codClick = "!*CLICK_" + StringUtil.strRight(texto, "!*CLICK_");
				codClick = StringUtil.strRight(codClick, "!*");
				codClick = StringUtil.strLeft(codClick, "*!");
				
				//pegando o documento do click com as informações
				doc = view.getDocumentByKey(codClick, true);

				if ( doc != null ) {
					/*
					Significado dos parametros:
					_inst = instalação
					_ccl = código do click
					_s = servidor
					_cp = referência da pessoa
					_cmo = referência do modelo de email
					_cme = referência da metrica de email
					 */
					String server, href;
					//String descricao = doc.getItemValueString("descricaoestatistica");
					server = (doc.getItemValueString("server")).replace("O=", "");
					server = server.replace("CN=", "");
					href = (doc.getItemValueString("host").toLowerCase()).replace("http://", ""); 
					href += "xtr/pagespublic.nsf/click_v1.xsp?";
					href += "&_inst=" + doc.getItemValueString("instalacao");;
					href += "&_ccl=" + 	doc.getItemValueString("xtr_cod");
					href += "&_s=" + server;
					href += "&_cp=" + docPessoa.getItemValueString("ref_pessoa");
					href += "&_cmo=" + docPessoa.getItemValueString("ref_modeloemail");
					href += "&_cme=" + docPessoa.getItemValueString("ref_metricaemail");
					texto = texto.replace( ("!*" + codClick + "*!"), href );
				}
			}
		} catch( NotesException e) {
			e.printStackTrace();
		}
		return texto;
	}

	/**
public void xtrRenderizaLink
Description: Essa função faz a busca do link no modelo de email e converte para o link principal
	 */
	public String xtrRenderizaComponente( String texto ) {
		try {
			Document docComp = null;
			String codComp, textoComp;
			Database dbComp = (Database) dbList.get("emailcomponente");
			View viewComp = dbComp.getView("EMAILCOMPONENTE-variavel");
			
			//resolvendo as variavel de email
			while ( texto.indexOf("!*COMPONENTE_") != -1 ) {
				codComp = "!*COMPONENTE_" + StringUtil.strRight(texto, "!*COMPONENTE_");
				codComp = StringUtil.strRight(codComp, "!*");
				codComp = StringUtil.strLeft(codComp, "*!");
				
				//se for um componente a variavel
				docComp = viewComp.getDocumentByKey(codComp, true);	
				
				if ( docComp != null ) {
					textoComp = returnTextOfMimeOrRichText(docComp, "corpo");
					texto = texto.replace( ("!*" + codComp + "*!"), textoComp );
				}
			}
		} catch ( NotesException e ) {
			e.printStackTrace();
		}
		return texto;
	}

	/**
	 * <strong>xtrRenderizaFormula</strong> - Resolve a formula caso houver algum erro retorna vazio o texto para ser substituido.
	 *@param texto - Texto que será replicado. 
	 */
	//@SuppressWarnings("unchecked")
	public String xtrRenderizaFormula( String texto ) {
		try {
			String formula;
			Vector<?> resultado;
			formula = texto.trim();
			formula = formula.replace('"', '\"');
			formula = formula.replace('{','"');
			formula = formula.replace('}', '"');
			resultado = session.evaluate(formula, executeDoc);
			return resultado.toString();
		} catch( Exception e) {
			e.printStackTrace();
			return texto;
		}
	}

	/**
	 * public void xtrRenderizaLink
	 * Description: Essa função faz a busca do link no modelo de email e converte para o link principal
	 */
	public String xtrRenderizaLink( String texto ) { 
		try {
			Document docLink = null;
			String codLink, link;
			String descricaoLink, hrefLink;
			Database dbLink = (Database) dbList.get("emaillink");
			View viewLink = dbLink.getView("EMAILLINK-variavel");
			
			//resolvendo as variavel de email
			while ( texto.indexOf("!*LINK_") != -1 ) {
				codLink = "!*LINK_" + StringUtil.strRight(texto, "!*LINK_");
				codLink = StringUtil.strRight(codLink, "!*");
				codLink = StringUtil.strLeft(codLink, "*!");

				//pegando o documento com as informações do link
				docLink = viewLink.getDocumentByKey(codLink, true);	

				if ( docLink != null ) {
					//descricaoEstatistica = docLink.descricaoEstatistica(0)
					descricaoLink = docLink.getItemValueString("descricaoLink");
					hrefLink = docLink.getItemValueString("link");
					link = "<a href=" + hrefLink + ">" + descricaoLink + "</a>";
					texto = texto.replace( ("!*" + codLink + "*!"), link );
				}
			}
		} catch( NotesException e ) {
			e.printStackTrace();
		}
		return texto;
	}


	/**
	public void xtrRenderizaOptout
	Description: Essa função renderiza a  variavel de email optout
	 */
	public String xtrRenderizaOptout ( String texto ) {
		try { 
			Database dbBlacklist;
			Database dbBlackCad;
			View viewBlacklistCad;
			Document docBCad;
			String dblcServer, dblcArquivo;
			String blServer, blArquivo;
			String textAnt,categ,desc,optemail;
			String optpagina, opttipo,ref,email,emails;
			String parametros;

			optpagina = opttipo = ref = email = emails = "";
			executeResult = "";
			
			if ( executeDoc != null ) { 
				dblcServer = executeDoc.getItemValueString("blacklistcadastro_servidor"); 
				dblcArquivo = executeDoc.getItemValueString("blacklistcadastro_arquivo");
				blServer = executeDoc.getItemValueString("blacklist_servidor");
				blArquivo = executeDoc.getItemValueString("blacklist_arquivo");
				dbBlackCad = session.getDatabase(dblcServer, dblcArquivo, false);
				viewBlacklistCad = dbBlackCad.getView("BLACKLISTCADASTRO-cod");

				if ( executeDoc.hasItem("blacklistxtrcod") ) {
					ref = executeDoc.getItemValueString("blacklistxtrcod");
					emails = "[";

					for( Object em : executeDoc.getItemValue("email") ) {
						email += "%22" + em.toString() + "%22,"; 	
					}

					emails += email.substring(0, (email.length() - 1) ) + "]";
					docBCad = viewBlacklistCad.getDocumentByKey(ref, true);

					textAnt = docBCad.getItemValueString("linktextoante");
					desc = docBCad.getItemValueString("linkdescricao");
					categ = docBCad.getItemValueString("categoria");
					categ = categ.replace(" ", "%20");
					optemail = docBCad.getItemValueString("optoutemail");
					optpagina = docBCad.getItemValueString("optoutpagina");
					opttipo = docBCad.getItemValueString("optouttipo");

					//parametros para criar o optout no email
					parametros = "{categoria:%22" + categ.trim() + "%22,";
					parametros += "email:" +emails.trim() + ",";
					parametros += "ref:%22" + ref.trim() + "%22,";
					parametros += "optemail:%22" + optemail + "%22,";
					parametros += "blservidor:%22" + blServer + "%22,";
					parametros += "blarquivo:%22" + blArquivo + "%22}";

					executeResult = textAnt.trim() + "<a href=\"" + optpagina.trim() + "?paarametros=" + parametros + "\">";
					executeResult += desc.trim() + "</a>";
					executeResult = executeResult.trim();
				}
			}
			return executeResult;
		} catch( NotesException e ) {
			e.printStackTrace();
			return texto;
		}
	}

}