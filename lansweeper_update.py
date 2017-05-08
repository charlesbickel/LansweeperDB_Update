import os
import pypyodbc


lsuser = NunYaBusiness...
lspwd = NunYaBusiness...


class Sweep:
    def __init__(self, asset, po, serial):

        self.asset = asset
        self.po = po
        self.serial = serial


        conn = pypyodbc.connect(
            r'DRIVER={SQL Server Native Client 11.0};'
            r'SERVER=NunYaBusiness;'
            r'DATABASE=lansweeperdb;'
            fr'UID={lsuser};'
            fr'PWD={lspwd};'
        )
        cursor = conn.cursor()

        try:
            cursor.execute(
                f"""
                UPDATE tblAssetCustom
                SET Custom1='{asset}', Custom2='{po}'
                WHERE Serialnumber='{serial}' AND (Custom1 IS NULL OR Custom2 IS NULL)
                """
            )

            conn.commit()

        except Exception as e:
            print(e)

        conn.close()
