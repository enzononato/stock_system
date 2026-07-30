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
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
