from linkedin_api import Linkedin
import pandas as pd

# A função é responsável por autenticar as contas de usuário para garantir acesso seguro ao LinkedIn.
def autenticar(email, senha):
    try:
        return Linkedin(email, senha)
    except Exception as e:
        print("Erro na Autenticação:", e)
        return None

# A função é responsável por realizar a leitura de arquivos Excel contendo informações dos usuários pesquisados no LinkedIn.
def LerExcel():
    try:
        UsuariosDf = pd.read_excel("Lista de Perfis.xlsx")
        return UsuariosDf
    except Exception as e:
        print("Ocorreu um erro ao ler o arquivo:", e)
        return None

# A função é responsável pelo armazenamento dos dados proveniente de arquivo Excel e pelo tratamento das variáveis dos usuários pesquisados no LinkedIn.
def UsandoExecel():
    try:
        usuarios = LerExcel()
        if usuarios is None:
            print("Erro ao ler o arquivo Excel")
            return None

        if "Usuarios" not in usuarios.columns:
            print("A coluna 'Usuarios' não foi encontrada no arquivo Excel")
            return None

        nomes_perfis = []  # Inicializa a lista de nomes de perfis
        for indice, usuario in enumerate(usuarios["Usuarios"]):
            partes = usuario.split("https://www.linkedin.com/in/")
            if len(partes) > 1:  # Verifica se há partes na URL
                nome_perfil = partes[1].split("/")[0]
                nomes_perfis.append(nome_perfil)  # Adiciona o nome do perfil à lista
        return nomes_perfis
    except Exception as e:
        print("Erro ao processar o arquivo Excel:", e)
        return None

# A função é responsável por extrair os dados pessoais dos usuários no LinkedIn.
def PesquisaDadosPessoais(api, nome):
    try:
        profile = api.get_profile(nome)  # Obter o perfil usando o nome fornecido
        dados_pessoais = {}
        if 'firstName' in profile and 'lastName' in profile:
            nome_completo = profile['firstName'] + " " + profile['lastName']
            dados_pessoais['Nome'] = nome_completo
        else:
            dados_pessoais['Nome'] = "Sem nome"
        return dados_pessoais
    except Exception as e:
        print(f"Erro ao pesquisar dados do nome candidato: {nome}: {e}")
        return {"Erro": "Erro ao pesquisar dados do nome candidato"}

# A função é responsável por extrair os dados de localização fornecidos pelos usuários ao LinkedIn.
def PesquisaLocalizacao(api, nome):
    try:
        profile = api.get_profile(nome)
        dados_localizacao = {}
        if 'geoLocationName' in profile:
            dados_localizacao['geoLocationName'] = profile['geoLocationName']
        else:
            dados_localizacao['geoLocationName'] = "Sem informações de localização"
        return dados_localizacao
    except Exception as e:
        print("Erro ao pesquisar localização: ", e)
        return {"Erro": "Erro ao pesquisar dados de localização do candidato"}

# A função é responsável por extrair os dados do último cargo fornecido pelos usuários ao LinkedIn.
def UltimoCargo(api, nome):
    try:
        profile = api.get_profile(nome)
        ultimo_cargo = "Desconhecido"
        if 'experience' in profile and len(profile['experience']) > 0:
            experiencias = profile['experience']
            experiencias.sort(key=lambda x: x.get('dateRange', {}).get('end', {}).get('year', 0), reverse=True)
            for experiencia in experiencias:
                if 'title' in experiencia:
                    ultimo_cargo = experiencia['title']
                    break
        return ultimo_cargo
    except Exception as e:
        print("Erro ao pesquisar último cargo: ", e)
        return {"Erro": "Erro ao pesquisar dados cargo do candidato"}

# A função é responsável por extrair os dados da empresa atual fornecida pelos usuários ao LinkedIn.
def PesquisaEmpresaAtual(api, nome):
    try:
        profile = api.get_profile(nome)
        if 'industryName' in profile:
            industry_name = profile['industryName']
            return industry_name
        else:
            return "Indústria não encontrada no perfil do usuário."
    except Exception as e:
        print("Erro ao pesquisar experiências:", e)
        return {"Erro": "Erro ao pesquisar empresa atual do candidato"}

# A função é responsável por extrair os dados das experiências fornecidas pelos usuários ao LinkedIn.
def PesquisaExperiencias(api, nome):
    try:
        profile = api.get_profile(nome)
        empresas = []
        if 'experience' in profile and len(profile['experience']) > 0:
            experiencias = profile['experience']
            for experiencia in experiencias:
                empresa = experiencia.get('companyName', 'Desconhecido')
                empresas.append(empresa)
        else:
            empresas.append("Sem informações de experiências")
        return ', '.join(empresas)
    except Exception as e:
        print("Erro ao pesquisar experiências: ", e)
        return {"Erro": "Erro ao pesquisar experiências do candidato"}

# A função é responsável por extrair os dados dos cursos realizados pelos usuários, fornecidos ao LinkedIn.
def PesquisaCursos(api, nome):
    try:
        profile = api.get_profile(nome)
        dados_cursos = []
        if 'education' in profile:
            cursos = profile['education']
            for curso in cursos:
                curso_info = []
                if 'schoolName' in curso:
                    curso_info.append(f"Nome da instituição: {curso['schoolName']}")
                if 'fieldOfStudy' in curso:
                    curso_info.append(f"Curso: {curso['fieldOfStudy']}")
                if 'degreeName' in curso:
                    curso_info.append(f"Grau: {curso['degreeName']}")
                dados_cursos.append(', '.join(curso_info))
        else:
            dados_cursos.append("Sem informações de cursos")

        return '\n'.join(dados_cursos)
    except Exception as e:
        print(f"Erro ao pesquisar cursos para o nome {nome}: {e}")
        return {"Erro": "Erro ao pesquisar cursos realizado pelo candidato"}

# A função é responsável por extrair os dados dos idiomas conhecidos pelo usuário no LinkedIn.
def PesquisaIdiomas(api, nome):
    try:
        profile = api.get_profile(nome)
        dados_idioma = {}
        if 'languages' in profile:
            idiomas = profile['languages']
            dados_idioma['Idiomas'] = []
            for idioma in idiomas:
                if 'name' in idioma and 'proficiency' in idioma:
                    nome_idioma = idioma['name']
                    proficiencia_idioma = idioma['proficiency']
                    dados_idioma['Idiomas'].append({'Idioma': nome_idioma, 'Proficiencia': proficiencia_idioma})
        else:
            dados_idioma['Idiomas'] = "Sem informações de idiomas"
        return dados_idioma
    except Exception as e:
        print("Erro ao pesquisar idiomas: ", e)
        return {"Erro": "Erro ao pesquisar idiomas conhecidos pelo candidato"}

# Função responsável por coordenar e executar as demais funções do script para realizar a pesquisa no LinkedIn
def UsandoAutomacao():
    usuario_index = 0
    consultas_realizadas = 0
    usuarios = [
        {"email": "kendrickabril1@gmail.com", "senha": "Kendrick@abril0"},
        #{"email": "researcher.experts@grupociadetalentos.com.br", "senha": "Experts22"},
    ]
    nomes = UsandoExecel()
    if nomes is None:
        print("Erro ao processar a lista de perfis")
        return

    total_usuarios = len(usuarios)
    total_nomes = len(nomes)
    resultados = []

    # Autentica com o primeiro usuário
    email = usuarios[usuario_index]["email"]
    senha = usuarios[usuario_index]["senha"]
    api = autenticar(email, senha)
    if api is None:
        print("Falha ao autenticar com o LinkedIn")
        return
    print(f"Autenticado como: {email}")

    while consultas_realizadas < total_nomes:
        # Alterna o usuário a cada 10 consultas
        if consultas_realizadas > 0 and consultas_realizadas % 5 == 0:
            usuario_index = (usuario_index + 1) % total_usuarios
            email = usuarios[usuario_index]["email"]
            senha = usuarios[usuario_index]["senha"]
            api = autenticar(email, senha)
            if api is None:
                print("Falha ao autenticar com o LinkedIn")
                return
            print(f"Autenticado como: {email}")

        # Pesquisa o próximo nome
        nome = nomes[consultas_realizadas]
        try:
            # Pesquisar e armazenar os dados pessoais para o usuário atual
            dados_pessoais = PesquisaDadosPessoais(api, nome)
            cargo_atual = UltimoCargo(api, nome)
            localizacao = PesquisaLocalizacao(api, nome)
            empresa = PesquisaEmpresaAtual(api, nome)
            experiencia = PesquisaExperiencias(api, nome)
            cursos = PesquisaCursos(api, nome)
            dados_idioma = PesquisaIdiomas(api, nome)

            idiomas = dados_idioma.get('Idiomas', [])
            if idiomas:
                idiomas_list = [idioma['Idioma'] for idioma in idiomas]
                proficiencias_list = [idioma['Proficiencia'] for idioma in idiomas]
                idiomas_str = ', '.join(idiomas_list)
                proficiencias_str = ', '.join(proficiencias_list)
            else:
                idiomas_str = ''
                proficiencias_str = ''

            resultado_usuario = {
                'Nome': dados_pessoais.get('Nome', ''),
                'Cargo Atual': cargo_atual,
                'Localizacao': localizacao.get('geoLocationName', ''),
                'Empresa Atual': empresa,
                'Experiências': experiencia,
                'Cursos': cursos,
                'Idiomas': idiomas_str,
                'Proficiência': proficiencias_str
            }
            resultados.append(resultado_usuario)

        except Exception as e:
            print(f"Erro ao pesquisar dados do nome candidato: {nome}: {e}")

        consultas_realizadas += 1

    df_resultados = pd.DataFrame(resultados)
    with pd.ExcelWriter('resultados_linkedin.xlsx') as writer:
        df_resultados.to_excel(writer, index=False)
    print("Dados de todos os usuários exportados para o arquivo 'resultados_linkedin.xlsx'")

if __name__ == "__main__":
    LerExcel()
    UsandoAutomacao()
