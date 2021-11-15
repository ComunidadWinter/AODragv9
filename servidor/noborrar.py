class AODrag:
    _public_methods_ = [ 'Test']
    _reg_progid_ = "AODrag"
    _reg_clsid_ = "{56173EAA-5C0C-4F73-A86D-598C494A18FD}"
    
    def Test(self, prueba):
        return prueba
    
if __name__ == "__main__":
    import win32com.server.register
    win32com.server.register.UseCommandLine(AODrag)
