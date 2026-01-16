# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy csproj and restore as distinct layers for better caching
COPY ["ConverterApi.csproj", "./"]
RUN dotnet restore "ConverterApi.csproj"

# Copy the rest of the files and build
COPY . .
RUN dotnet build "ConverterApi.csproj" -c Release -o /app/build

# Publish the application
FROM build AS publish
RUN dotnet publish "ConverterApi.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app

# Install system dependencies for PdfSharpCore (libgdiplus)
RUN apt-get update && apt-get install -y \
    libgdiplus \
    libc6-dev \
    fontconfig \
    fonts-liberation \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

COPY --from=publish /app/publish .

# Render typically provides a PORT environment variable
ENV ASPNETCORE_URLS=http://0.0.0.0:8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "ConverterApi.dll"]

