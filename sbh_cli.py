import click
from utils import *
from main.models import PurchaseOrder, Resource
from sbh_creator import SBH


def get_contractor():
    contractors = [
        # (Choice #, CONTRACTOR)
        (1, 'AL ROSTAMANI'),
        (2, 'CANAL'),
        (3, 'INNOVATION'),
        (4, 'INTELTEC'),
        (5, 'PENTA'),
        (6, 'REACH'),
        (7, 'SGEM'),
        (8, 'SHAHID'),
        (9, 'SKYLOG'),
        (10, 'STAR SERVICES'),
        (11, 'TAMDEED'),
        (12, 'TASC'),
        (13, 'TECHNOLIGIA'),
        (14, 'TELEPHONY'),
        (15, 'XAD TECHNOLOGIES'),
    ]

    click.echo("Please select a CONTRACTOR:")
    for row in contractors:
        click.echo("  {}: {}".format(row[0], row[1]))

    row_num = click.prompt("Enter your choice: ", type=int)
    return contractors[row_num - 1]


def get_cycle():
    cycles = [
        # (Choice #, period_start, period_end, ramadan_start, ramdan_end, cycle)
        (1, '20/01/2016', '19/02/2016', None, None),
        (2, '20/02/2016', '19/03/2016', None, None),
        (3, '20/03/2016', '19/04/2016', None, None),
        (4, '20/04/2016', '19/05/2016', None, None),
        (5, '20/05/2016', '19/06/2016', '5/6/2016', '19/06/2016'),
        (6, '20/06/2016', '19/07/2016', '20/06/2016', '08/07/2016'),
        (7, '20/07/2016', '19/08/2016', None, None),
        (8, '20/08/2016', '19/09/2016', None, None),
        (9, '20/09/2016', '19/10/2016', None, None),
        (10, '20/10/2016', '19/11/2016', None, None),
        (11, '20/11/2016', '19/12/2016', None, None),
        (12, '20/12/2016', '19/01/2017', None, None)
    ]

    click.echo("Please select a CYCLE:")
    for row in cycles:
        click.echo("  {}: cycle {}-{} | ramadan {}-{}".format(row[0], row[1], row[2],
                                                              row[3] or '', row[4] or '')
                   )
    row_num = click.prompt("Enter your choice: ", type=int)

    return cycles[row_num - 1]


def get_division(po_num):
    qs_po = PurchaseOrder.objects.filter(po_num=po_num).first()
    qs_resource = Resource.objects.filter(po_id=qs_po.id)
    divisions = list(qs_resource.values_list('division', flat=True).distinct())
    click.echo("Please select a DIVISION:")
    for row in divisions:
        click.echo("  {}: {}".format(divisions.index(row), row))

    row_num = click.prompt("Enter your choice: ", type=int)

    return divisions[row_num]

@click.command()
def cli():
    sbh = SBH()
    click.echo("#****************************************#")
    click.echo("#----------------------------------------#")
    click.echo("# Technology Budget Planning and Control #")
    click.echo("# SBH FORM CREATOR v0.1                  #")
    click.echo("#----------------------------------------#")
    click.echo("#****************************************#")
    click.echo("")
    while True:
        cycle = get_cycle()
        click.echo("Do you want to generate for,")
        click.echo("  1. ALL")
        click.echo("  2. Per PO NUMBER")
        click.echo("  3. Per CONTRACTOR")
        click.echo("  4. Per DIVISION")
        generate_category = click.prompt("Enter your choice: ", type=int, default=1)

        if generate_category == 2:
            po_num = click.prompt("Enter your PO NUMBER: ", type=str)
            click.echo("Generating for {}".format(po_num))
            sbh.make_sbh_per_po(po_num, cycle[1], cycle[2], cycle[3], cycle[4])
            click.echo("DONE")

        elif generate_category == 3:
            contractor = get_contractor()
            click.echo("Generating for {}".format(contractor[1]))
            sbh.make_sbh_per_contractor(contractor[1], cycle[1], cycle[2], cycle[3], cycle[4])
            click.echo("DONE")

        elif generate_category == 4:
            po_num = click.prompt("Enter your PO NUMBER: ", type=str)
            division = get_division(po_num)
            click.echo("Generating for {}".format(po_num))
            sbh.make_sbh_per_division(po_num, division, cycle[1], cycle[2], cycle[3], cycle[4])
            click.echo("DONE")

        else:
            contractor = 'all'
            click.echo("Generating for {}".format(contractor))
            sbh.make_sbh_per_contractor(contractor, cycle[1], cycle[2])
            click.echo("DONE")

        click.confirm('Do you want to continue?', abort=True)


if __name__ == '__main__':
    cli()
