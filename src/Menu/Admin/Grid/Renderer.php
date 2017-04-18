<?php

namespace CubeSystems\Leaf\Menu\Admin\Grid;

use CubeSystems\Leaf\Admin\Grid;
use CubeSystems\Leaf\Admin\Layout\Footer;
use CubeSystems\Leaf\Admin\Layout\Footer\Tools;
use CubeSystems\Leaf\Admin\Widgets\Link;
use CubeSystems\Leaf\Html\Elements\Content;
use CubeSystems\Leaf\Html\Elements\Element;
use CubeSystems\Leaf\Html\Html;
use CubeSystems\Leaf\Menu\AbstractItem;
use CubeSystems\Leaf\Nodes\MenuItem;
use Illuminate\Contracts\Pagination\Paginator;
use Illuminate\Support\Collection;

class Renderer
{
    /**
     * @var Grid
     */
    protected $grid;

    /**
     * @var Paginator
     */
    protected $page;

    /**
     * Renderer constructor.
     * @param Grid $grid
     */
    public function __construct( Grid $grid )
    {
        $this->grid = $grid;
    }

    /**
     * @return Grid
     */
    public function grid()
    {
        return $this->grid;
    }

    /**
     * @return Content
     */
    protected function table()
    {
        return new Content( [
            Html::header( [
                Html::h1( trans( 'leaf::resources.all_resources' ) ),
                Html::span( trans( 'leaf::pagination.items_found', [ 'total' => $this->page->total() ] ) )
                    ->addClass( 'extras totals only-text' )
            ] ),
            Html::div(
                Html::div(
                    $this->buildTree( $this->page->getCollection(), 1 )
                )->addClass( 'collection' )
            )->addClass( 'body' )
        ] );
    }

    /**
     * @param Collection $items
     * @param int $level
     * @return Element
     */
    protected function buildTree( Collection $items, $level = 1 )
    {
        $url = $this->url( 'edit', '__ID__' );

        $list = Html::ul()->addAttributes( [ 'data-level' => $level ] );

        $this->reorderItems( $items );

        foreach( $items as $item )
        {
            $children = $item->children;
            $hasChildren = ( $children && $children->count() );

            $li = Html::li()
                ->addAttributes( [
                    'data-level' => $level,
                    'data-id' => $item->getKey(),
                ] );

            $cell = Html::div()->addClass( 'node-cell active' );

            $link = str_replace( '__ID__', $item->getKey(), $url );

            foreach( $this->grid()->getColumns() as $column )
            {
                $cell->append( Html::link(
                    Html::span( $item->{$column->getName()} )
                )
                    ->addAttributes( [ 'href' => $link ] )
                );
            }

            $li->append( $cell );

            if( $hasChildren )
            {
                $li->append( $this->buildTree( $children, $level + 1 ) );
            }

            if( $level === 1 && $item->hasParent() )
            {
                continue;
            }

            $list->append( $li );
        }

        return $list;
    }

    /**
     * @param Collection $items
     * @return void
     */
    protected function reorderItems( Collection $items )
    {
        foreach( $items as $model )
        {
            /** @var MenuItem $model */
            if( !$model->isAfter() )
            {
                continue;
            }

            $afterItem = $items->filter( function( MenuItem $item ) use ( $model )
            {
                return $item->getId() === $model->getAfterId();
            } )->first();

            $currentPosition = $items->search( $model );
            $afterKey = $items->search( $afterItem );

            $items->forget( $currentPosition );

            $items->splice( ++$afterKey, 0, [ $model ] );
        }
    }

    /**
     * @return Element
     */
    protected function footer()
    {
        $createButton = Link::create( $this->url( 'create' ) )
            ->asButton( 'primary' )
            ->withIcon( 'plus' )
            ->title( trans( 'leaf::resources.create_new' ) );

        $tools = new Tools();
        $tools->getBlock( 'primary' )->push( $createButton );

        $footer = new Footer( 'main' );
        $footer->getRows()->prepend( $tools );

        return $footer->render();
    }

    /**
     * @param $route
     * @param array $parameters
     * @return \CubeSystems\Leaf\Admin\Module\Route
     */
    public function url( $route, $parameters = [] )
    {
        return $this->grid()->getModule()->url( $route, $parameters );
    }

    /**
     * @param Paginator $page
     * @return Element
     */
    public function render( Paginator $page )
    {
        $this->page = $page;

        return Html::section( [
            $this->table(),
            $this->footer(),
        ] );
    }
}