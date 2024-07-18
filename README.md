![val.png](media/val.png)
> # SISTEMA VAL
> Aplicação para automação de valoração de ordens de
> serviços executados pela contratada via *SAPGUI*.

## Instalação
1. Clone o repositório
2. Instale as dependências
```bash
pip install -r requirements.txt
```

## Utilização
1. Esteja dentro do ambiente SABESP.
2. Tenha um login com acesso ao *SAPGUI*
3. Execute o arquivo `main.py`
```bash
python -m src.main
```
## Uso de 6 sessões simultâneas
Para executar 6 tipos diferentes de famílias, executar via compilador do rust.
```bash
cargo run
```
**Observação:** Edite `main.rs` para ajustar as famílias. É obrigatório ter  instalado **rustup e rustc** na máquina.

### Execução
1. Preencha os campos obrigatórios
2. Selecione o tipo de pagamento, contratada tipo serviços
3. Dê o Enter para inicializar
4. Aguarde o processamento

## Desenvolvimento
Desenvolvido por Hyslan Silva Cruz, hyslansilva@gmail.com
