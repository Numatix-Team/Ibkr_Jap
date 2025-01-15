from ib_insync import ContFuture

def create_cont_future(symbol, exchange, localSymbol='', multiplier='', currency='USD'):
    return ContFuture(
        symbol=symbol,
        exchange=exchange,
        localSymbol=localSymbol,
        multiplier=multiplier,
        currency=currency
    )

def cont_future_to_dict(cont_future):
    return cont_future.dict()

def update_cont_future(cont_future, **kwargs):
    cont_future.update(**kwargs)
    return cont_future

def get_non_defaults(cont_future):
    return cont_future.nonDefaults()

def cont_future_to_tuple(cont_future):
    return cont_future.tuple()

# Example Usage
if __name__ == "__main__":
    contract = create_cont_future(
        symbol='N225M',
        exchange='OSE.JPN',
        localSymbol='N225M_CONT',
        multiplier='1',
        currency='JPY'
    )
    print("Contract as dict:", cont_future_to_dict(contract))

    updated_contract = update_cont_future(contract, symbol='N225M_NEW')
    print("Updated Contract:", cont_future_to_dict(updated_contract))

    non_defaults = get_non_defaults(contract)
    print("Non-default fields:", non_defaults)

    contract_tuple = cont_future_to_tuple(contract)
    print("Contract as tuple:", contract_tuple)