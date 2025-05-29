import cartopy.crs as ccrs
import cartopy.feature as cfeature
import matplotlib.pyplot as plt
import utm

# Lista de pontos UTM (zona 22 sul)


def mapa_qag(df_pontos):

    linhas = []
    for linha in df_pontos.iterrows():
        linhas.append(linha[1].to_list())

    print(linhas)

    pontos_geo = [
        utm.to_latlon(e, n, 22, northern=False) + (nome,) for nome, e, n in linhas
    ]

    # Criar mapa
    fig = plt.figure(figsize=(10, 10))
    ax = fig.add_subplot(1, 1, 1, projection=ccrs.PlateCarree())
    ax.set_title("Pontos PSFS - Santa Catarina", fontsize=14)

    ax.add_feature(cfeature.LAND)
    ax.add_feature(cfeature.OCEAN)
    ax.add_feature(cfeature.COASTLINE)
    ax.add_feature(cfeature.BORDERS, linestyle=":")
    ax.add_feature(cfeature.LAKES, alpha=0.5)
    ax.add_feature(cfeature.RIVERS)

    # Zoom no litoral norte de SC
    ax.set_extent([-48.7, -48.3, -26.5, -26.0], crs=ccrs.PlateCarree())
    for ponto_geo in pontos_geo:
        for nome, lat, lon in pontos_geo:
            ax.plot(
                lon,
                lat,
                marker="o",
                color="red",
                markersize=4,
                transform=ccrs.PlateCarree(),
            )
            ax.text(
                lon + 0.01, lat + 0.01, nome, fontsize=6, transform=ccrs.PlateCarree()
            )

        plt.tight_layout()
        plt.show()
