
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import json

def obtener_ruta_archivo():
    ruta = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
    if ruta:
        return ruta
    else:
        print("No se ha seleccionado ningún archivo")
        exit()

# Función para leer el archivo Excel y generar el JSON
def generar_json(ruta_archivo):
    # Leer el archivo Excel
    datos_excel = pd.read_excel(ruta_archivo)

    # Crear el diccionario JSON
    json_data = {
        "numDocumentoIdObligado": str(datos_excel["numDocumentoIdObligado"][0]),
        "numFactura": str(datos_excel["numFactura"][0]),
        "TipoNota": str(datos_excel["TipoNota"][0]),
        "numNota": str(datos_excel["numNota"][0]),
        "usuarios": [
            {
                "tipoDocumentoIdentificacion": str(datos_excel["tipoDocumentoIdentificacion"][0]),
                "numDocumentoIdentificacion": str(datos_excel["numDocumentoIdentificacion"][0]),
                "tipoUsuario": str(datos_excel["tipoUsuario"][0]),
                "fechaNacimiento": str(datos_excel["fechaNacimiento"][0]),
                "codSexo": str(datos_excel["codSexo"][0]),
                "codPaisResidencia": str(datos_excel["codPaisResidencia"][0]),
                "codMunicipioResidencia": str(datos_excel["codMunicipioResidencia"][0]),
                "codZonaTerritorialResidencia": str(datos_excel["codZonaTerritorialResidencia"][0]),
                "incapacidad": str(datos_excel["incapacidad"][0]),
                "consecutivo": str(datos_excel["consecutivo"][0]),
                "codPaisOrigen": str(datos_excel["codPaisOrigen"][0]),
                "servicios": {
                    "consultas": [
                        {
                            "codPrestador": "",
                            "fechaInicioAtencion": "",
                            "numAutorizacion": "",
                            "codConsulta": "",
                            "modalidadGrupoServicioTecSal": "",
                            "grupoServicios": "",
                            "codServicio": "",
                            "finalidadTecnologiaSalud": "",
                            "causaMotivoAtencion": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoRelacionado1": "",
                            "codDiagnosticoRelacionado2": "null",
                            "codDiagnosticoRelacionado3": "null",
                            "tipoDiagnosticoPrincipal": "",
                            "tipoDocumentoIdentificacion": "",
                            "numDocumentoIdentificacion": "",
                            "vrServicio": "",
                            "conceptoRecaudo": "",
                            "valorPagoModerador": "",
                            "numFEVPagoModerador": "",
                            "consecutivo": ""
                        },
                        {
                            "codPrestador": "",
                            "fechaInicioAtencion": "",
                            "numAutorizacion": "",
                            "codConsulta": "",
                            "modalidadGrupoServicioTecSal": "",
                            "grupoServicios": "",
                            "codServicio": "",
                            "finalidadTecnologiaSalud": "",
                            "causaMotivoAtencion": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoRelacionado1": "",
                            "codDiagnosticoRelacionado2": "null",
                            "codDiagnosticoRelacionado3": "null",
                            "tipoDiagnosticoPrincipal": "",
                            "tipoDocumentoIdentificacion": "",
                            "numDocumentoIdentificacion": "",
                            "vrServicio": "",
                            "conceptoRecaudo": "",
                            "valorPagoModerador": "",
                            "numFEVPagoModerador": "",
                            "consecutivo": ""
                        },
                    ],
                    "procedimientos": [
                        {
                            "codPrestador": "",
                            "fechaInicioAtencion": "",
                            "idMIPRES": "null",
                            "numAutorizacion": "null",
                            "codProcedimiento": "",
                            "viaIngresoServicioSalud": "",
                            "modalidadGrupoServicioTecSal": "",
                            "grupoServicios": "",
                            "codServicio": "",
                            "finalidadTecnologiaSalud": "",
                            "tipoDocumentoIdentificacion": "",
                            "numDocumentoIdentificacion": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoRelacionado": "",
                            "codComplicacion": "",
                            "vrServicio": "",
                            "conceptoRecaudo": "",
                            "valorPagoModerador": "",
                            "numFEVPagoModerador": "",
                            "consecutivo": ""
                        },
                    ],
                    "urgencias": [
                        {
                            "codPrestador": "",
                            "fechaInicioAtencion": "",
                            "causaMotivoAtencion": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoPrincipalE": "",
                            "codDiagnosticoRelacionadoE1": "null",
                            "codDiagnosticoRelacionadoE2": "null",
                            "codDiagnosticoRelacionadoE3": "null",
                            "condicionDestino": "",
                            "codDiagnosticoCausaMuerte": "null",
                            "fechaEgreso": "",
                            "consecutivo": ""
                        },
                    ],
                    "hospitalizacion": [
                        {
                            "codPrestador": "",
                            "viaIngresoServicioSalud": "",
                            "fechaInicioAtencion": "",
                            "numAutorizacion": "",
                            "causaMotivoAtencion": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoPrincipalE": "",
                            "codDiagnosticoRelacionadoE1": "null",
                            "codDiagnosticoRelacionadoE2": "null",
                            "codDiagnosticoRelacionadoE3": "null",
                            "codComplicacion": "null",
                            "condicionDestinoUsuarioEgreso": "",
                            "codDiagnosticoCausaMuerte": "null",
                            "fechaEgreso": "",
                            "consecutivo": ""
                        },
                    ],
                    "recienNacidos": [
                        {
                            "codPrestador": "",
                            "tipoDocumentoIdentificacion": "",
                            "numDocumentoIdentificacion": "",
                            "fechaNacimiento": "",
                            "edadGestacional": "",
                            "numConsultasCPrenatal": "",
                            "codSexoBiologico": "",
                            "peso": "",
                            "codDiagnosticoPrincipal": "",
                            "condicionDestinoUsuarioEgreso": "",
                            "codDiagnosticoCausaMuerte": "",
                            "fechaEgreso": "",
                            "consecutivo": ""
                        },
                    ],
                    "medicamentos": [
                        {
                            "codPrestador": "",
                            "numAutorizacion": "",
                            "idMIPRES": "",
                            "fechaDispensAdmon": "",
                            "codDiagnosticoPrincipal": "",
                            "codDiagnosticoRelacionado": "",
                            "tipoMedicamento": "",
                            "codTecnologiaSalud": "",
                            "nomTecnologiaSalud": "null",
                            "concentracionMedicamento": "",
                            "unidadMedida": "",
                            "formaFarmaceutica": "",
                            "unidadMinDispensa": "",
                            "cantidadMedicamento": "",
                            "diasTratamiento": "",
                            "tipoDocumentoIdentificacion": "",
                            "numDocumentoIdentificacion": "",
                            "vrUnitMedicamento": "",
                            "vrServicio": "",
                            "conceptoRecaudo": "",
                            "valorPagoModerador": "",
                            "numFEVPagoModerador": "",
                            "consecutivo": ""
                        },
                    ],
                    "otrosServicios": [
                        {
                            "codPrestador": str(datos_excel["codPrestador"][0]),
                            "numAutorizacion": str(datos_excel["numAutorizacion"][0]),
                            "idMIPRES": str(datos_excel["idMIPRES"][0]),
                            "fechaSuministroTecnologia": str(datos_excel["fechaSuministroTecnologia"][0]),
                            "tipoOS": str(datos_excel["tipoOS"][0]),
                            "codTecnologiaSalud": str(datos_excel["codTecnologiaSalud"][0]),
                            "nomTecnologiaSalud": str(datos_excel["nomTecnologiaSalud"][0]),
                            "cantidadOS": str(datos_excel["cantidadOS"][0]),
                            "tipoDocumentoIdentificacion": str(datos_excel["tipoDocumentoIdentificacionOtrosServicios"][0]),
                            "numDocumentoIdentificacion": str(datos_excel["numDocumentoIdentificacionOtrosServicios"][0]),
                            "vrUnitOS": str(datos_excel["vrUnitOS"][0]),
                            "vrServicio": str(datos_excel["vrServicio"][0]),
                            "conceptoRecaudo": str(datos_excel["conceptoRecaudo"][0]),
                            "valorPagoModerador": str(datos_excel["valorPagoModerador"][0]),
                            "numFEVPagoModerador": str(datos_excel["numFEVPagoModerador"][0]),
                            "consecutivo": str(datos_excel["consecutivoOtrosServicios"][0])
                        },
                    ],
                },
            },
        ],
    }

    # Specify the file path where you want to save the JSON file
    file_path = "example.json"

    # Write the data to the JSON file
    with open(file_path, 'w') as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f"JSON data has been saved to {file_path}")

def on_button1_click():
    ruta_archivo = obtener_ruta_archivo()
    print("IMPRIMIENDO...")
    print(ruta_archivo)
    if ruta_archivo:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, str(ruta_archivo))
        entry_path.config(state='readonly')
        entry_path.config(width= len(ruta_archivo) +10 )
        #generar_json(ruta_archivo)

def on_button2_click():
    print("Botón 2 clickeado")
    generar_json(entry_path.get())

def salir_pantalla_completa(event=None):
        root.attributes('-fullscreen', False)
        root.quit()

   

# Configuración de la ventana principal
root = tk.Tk()
#root.attributes('-fullscreen',True)
 # Salir de la pantalla completa con un clic en un botón
#button = tk.Button(root, text="Salir de pantalla completa", command=salir_pantalla_completa)
#button.pack()
root.resizable(height = None, width = None)
root.geometry("800x600")

root.title("Generar JSON desde Excel")




# Configuración de los botones
button1 = tk.Button(root, text="Seleccionar archivo Excel", command=on_button1_click)
button1.pack(pady=5)


# Campo de texto para mostrar la ruta del archivo seleccionado
entry_path = tk.Entry(root)
entry_path.pack(pady=5)

button2 = tk.Button(root, text="Generar JSON", command=on_button2_click)
button2.pack(pady=5)

# Ejecutar el bucle principal de la interfaz gráfica
root.mainloop()