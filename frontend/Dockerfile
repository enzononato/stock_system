# Etapa 1: build
#
# A versão de Node precisa ser a mesma que gerou o package-lock.json. Com
# node:20 o `npm ci` falhava: o npm 10 resolve a árvore de dependências de
# forma diferente do npm que acompanha o Node 24 e reclamava de pacotes
# "Missing from lock file" (esbuild e os binários @esbuild/*) que a versão nova
# nem considera necessários. O lock só é confiável para o npm que o produziu.
# Além disso, @testing-library/jest-dom exige Node >= 22.
# O campo "engines" do package.json documenta esse requisito.
FROM node:24-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build

# Etapa 2: servidor Nginx
FROM nginx:alpine
COPY --from=build /app/dist /usr/share/nginx/html

# Template (não config final): a imagem oficial roda envsubst sobre todo
# arquivo em /etc/nginx/templates/*.template na inicialização, gravando o
# resultado em /etc/nginx/conf.d/ com o mesmo nome sem o sufixo .template.
# Só substitui as variáveis realmente definidas no ambiente do container, e
# nunca as variáveis nativas do nginx ($host, $remote_addr etc.), que usam a
# mesma sintaxe $NOME — por isso é seguro usar num arquivo cheio delas.
#
# Os defaults abaixo reproduzem o nome/porta do serviço no docker-compose.yml
# deste repo, então `docker compose up` continua funcionando sem nenhuma
# variável extra. Numa plataforma que não permita nomear o serviço do backend
# como "backend" (ex.: EasyPanel, Railway), sobrescreva BACKEND_HOST (e
# BACKEND_PORT, se a porta interna não for 8000) nas variáveis de ambiente do
# serviço do frontend — sem precisar tocar nesta imagem.
COPY nginx.conf.template /etc/nginx/templates/default.conf.template
ENV BACKEND_HOST=backend
ENV BACKEND_PORT=8000

EXPOSE 80
